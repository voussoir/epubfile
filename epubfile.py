import copy
import functools
import io
import mimetypes
import os
import re
import tempfile
import urllib.parse
import uuid
import zipfile

import bs4
import tinycss2

from voussoirkit import getpermission
from voussoirkit import pathclass

HTML_LINK_PROPERTIES = {
    'a': ['href'],
    'audio': ['src'],
    'image': ['href', 'xlink:href'],
    'img': ['src'],
    'link': ['href'],
    'script': ['src'],
    'source': ['src'],
    'track': ['src'],
    'video': ['src', 'poster'],
}

EXTENSION_MIMETYPES = {
    'htm': 'application/xhtml+xml',
    'html': 'application/xhtml+xml',
    'otf': 'font/otf',
    'pls': 'application/pls+xml',
    'smi': 'application/smil+xml',
    'smil': 'application/smil+xml',
    'sml': 'application/smil+xml',
    'ttf': 'font/ttf',
    'woff': 'font/woff',
    'woff2': 'font/woff2',
    'xhtml': 'application/xhtml+xml',
    'xpgt': 'application/vnd.adobe-page-template+xml',
}

MIMETYPE_DIRECTORIES = {
    'application/font-sfnt': 'Fonts',
    'application/x-dtbncx+xml': '.',
    'application/x-font-ttf': 'Fonts',
    'application/xhtml+xml': 'Text',
    'audio': 'Audio',
    'font': 'Fonts',
    'image': 'Images',
    'text/css': 'Styles',
    'video': 'Video',
}

MIMETYPE_FILE_TEMPLATE = 'application/epub+zip'

CONTAINER_XML_TEMPLATE = '''
<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
    <rootfiles>
        <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
   </rootfiles>
</container>
'''.strip()

OPF_TEMPLATE = '''
<?xml version="1.0" encoding="utf-8"?>
<package version="3.0" unique-identifier="BookId" xmlns="http://www.idpf.org/2007/opf">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:identifier id="BookId">{uuid}</dc:identifier>
    <dc:creator id="cre">author</dc:creator>
    <meta scheme="marc:relators" refines="#cre" property="role">aut</meta>
    <dc:title>title</dc:title>
    <dc:language>und</dc:language>
  </metadata>
  <manifest>
    <item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>
    <item id="nav.xhtml" href="Text/nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>
  </manifest>
  <spine toc="ncx">
    <itemref idref="nav.xhtml" linear="no"/>
  </spine>
</package>
'''.strip()

NCX_TEMPLATE = '''
<?xml version="1.0" encoding="utf-8"?>
<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">
  <head>
    <meta name="dtb:uid" content="{uuid}" />
  </head>
<docTitle>
   <text>{title}</text>
</docTitle>
<navMap>
{navpoints}
</navMap>
</ncx>
'''.strip()

NAV_XHTML_TEMPLATE = '''
<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops">
<head>
  <meta charset="utf-8"/>
</head>
<body epub:type="frontmatter">
  <nav epub:type="toc" id="toc">
    <h1>Table of Contents</h1>
    {toc_contents}
  </nav>
</body>
</html>
'''.strip()

TEXT_TEMPLATE = '''
<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops">
<head>
{head_content}
</head>
<body>
{body_content}
</body>
</html>
'''.strip()


# EPUB COMPRESSION
################################################################################
def compress_epub(directory, epub_filepath):
    directory = pathclass.Path(directory)
    epub_filepath = pathclass.Path(epub_filepath)

    if epub_filepath in directory:
        raise ValueError('Epub inside its own directory')

    if epub_filepath.extension != 'epub':
        epub_filepath = epub_filepath.add_extension('epub')

    with zipfile.ZipFile(epub_filepath.absolute_path, 'w') as z:
        z.write(directory.with_child('mimetype').absolute_path, arcname='mimetype')
        for file in directory.walk():
            if file in [directory.with_child('mimetype'), directory.with_child('sigil.cfg')]:
                continue
            z.write(
                file.absolute_path,
                arcname=file.relative_to(directory),
                compress_type=zipfile.ZIP_DEFLATED,
            )
    return epub_filepath

def extract_epub(epub_filepath, directory):
    epub_filepath = pathclass.Path(epub_filepath)
    directory = pathclass.Path(directory)

    with zipfile.ZipFile(epub_filepath.absolute_path, 'r') as z:
        z.extractall(directory.absolute_path)

# XHTML TOOLS
################################################################################
def fix_xhtml(xhtml, return_soup=False):
    if isinstance(xhtml, bs4.BeautifulSoup):
        soup = xhtml
    else:
        # For the text pages, html5lib is the best because html.parser and lxml
        # lowercase all attributes, breaking svg's case-sensitive viewBox etc.
        # and xml loses all of the namespaces when namespaced elements are nested
        # like <html xmlns="..."><svg xmlns:xlink="..."></svg></html>.
        # The downside of html5lib is it turns the xml declaration at the top
        # into a comment which we must undo manually.
        soup = bs4.BeautifulSoup(xhtml, 'html5lib')

    if not soup.html:
        html = soup.new_tag('html')
        for child in list(soup.contents):
            html.append(child)
        soup.append(html)

    if not soup.html.body:
        body = soup.new_tag('body')
        for child in list(soup.html.contents):
            body.append(child)
        soup.html.append(body)

    if not soup.html.get('xmlns'):
        soup.html['xmlns'] = 'http://www.w3.org/1999/xhtml'

    try:
        doctype = next(i for i in soup.contents if isinstance(i, bs4.Doctype))
    except StopIteration:
        doctype = bs4.Doctype('html')
        soup.html.insert_before(doctype)

    # html5lib turns the xml declaration into a comment which we must revert.
    try:
        if isinstance(soup.contents[0], bs4.Comment):
            declaration = bs4.Declaration('xml version="1.0" encoding="utf-8"')
            soup.insert(0, declaration)
            declaration.next.extract()
    except StopIteration:
        pass

    try:
        declaration = next(i for i in soup.contents if isinstance(i, bs4.Declaration))
    except StopIteration:
        declaration = bs4.Declaration('xml version="1.0" encoding="utf-8"')
        doctype.insert_before(declaration)

    if return_soup:
        return soup
    return str(soup)

def xhtml_replacements(xhtml, replacements, return_soup=False):
    if isinstance(xhtml, bs4.BeautifulSoup):
        xhtml = str(xhtml)

    for (re_from, re_to) in replacements:
        xhtml = re.sub(re_from, re_to, xhtml, flags=re.DOTALL)

    if return_soup:
        soup = bs4.BeautifulSoup(xhtml, 'html5lib')
        return soup

    return xhtml

def demote_xhtml_headers(xhtml, return_soup=False):
    replacements = [
        (r'<h5([^>]*?>.*?)</h5>', r'<h6\1</h6>'),
        (r'<h4([^>]*?>.*?)</h4>', r'<h5\1</h5>'),
        (r'<h3([^>]*?>.*?)</h3>', r'<h4\1</h4>'),
        (r'<h2([^>]*?>.*?)</h2>', r'<h3\1</h3>'),
        (r'<h1([^>]*?>.*?)</h1>', r'<h2\1</h2>'),
    ]
    return xhtml_replacements(xhtml, replacements, return_soup=return_soup)

def promote_xhtml_headers(xhtml, return_soup=False):
    replacements = [
        (r'<h2([^>]*?>.*?)</h2>', r'<h1\1</h1>'),
        (r'<h3([^>]*?>.*?)</h3>', r'<h2\1</h2>'),
        (r'<h4([^>]*?>.*?)</h4>', r'<h3\1</h3>'),
        (r'<h5([^>]*?>.*?)</h5>', r'<h4\1</h4>'),
        (r'<h6([^>]*?>.*?)</h6>', r'<h5\1</h5>'),
    ]
    return xhtml_replacements(xhtml, replacements, return_soup=return_soup)

# MIMETYPE DECISIONMAKERS
################################################################################
def get_directory_for_mimetype(mime):
    directory = (
        MIMETYPE_DIRECTORIES.get(mime) or
        MIMETYPE_DIRECTORIES.get(mime.split('/')[0]) or
        'Misc'
    )
    return directory

def get_mimetype_for_basename(basename):
    extension = os.path.splitext(basename)[1].strip('.')
    mime = (
        EXTENSION_MIMETYPES.get(extension) or
        mimetypes.guess_type(basename)[0] or
        'application/octet-stream'
    )
    return mime

# OPF ELEMENT GENERATORS
################################################################################
def make_manifest_item(id, href, mime):
    manifest_item = f'<item id="{id}" href="{href}" media-type="{mime}"/>'
    # 'html.parser' just for having the simplest output.
    manifest_item = bs4.BeautifulSoup(manifest_item, 'html.parser')
    return manifest_item.item

def make_meta_item(content=None, attrs=None):
    if content:
        meta_item = f'<meta>{content}</meta>'
    else:
        meta_item = f'<meta/>'
    # 'html.parser' just for having the simplest output.
    meta_item = bs4.BeautifulSoup(meta_item, 'html.parser')
    if attrs:
        meta_item.attrs.update(attrs)
    return meta_item.meta

def make_spine_item(id):
    spine_item = f'<itemref idref="{id}"/>'
    # 'html.parser' just for having the simplest output.
    spine_item = bs4.BeautifulSoup(spine_item, 'html.parser')
    return spine_item.itemref

# DECORATORS
################################################################################
def writes(method):
    @functools.wraps(method)
    def wrapped_method(self, *args, **kwargs):
        if self.read_only:
            raise ReadOnly(method.__qualname__)
        return method(self, *args, **kwargs)
    return wrapped_method

# CLASSES
################################################################################
class EpubfileException(Exception):
    error_message = ''

    def __init__(self, *args, **kwargs):
        super().__init__()
        self.given_args = args
        self.given_kwargs = kwargs
        self.error_message = self.error_message.format(*args, **kwargs)
        self.args = (self.error_message, args, kwargs)

    def __str__(self):
        return self.error_message

class InvalidEpub(EpubfileException):
    error_message = '{} is invalid: {}'

class FileExists(EpubfileException):
    error_message = 'There is already a file at {}.'

class IDExists(EpubfileException):
    error_message = 'There is already a file with id {}.'

class NotInManifest(EpubfileException):
    error_message = '{} is not in the manifest.'

class NotInSpine(EpubfileException):
    error_message = '{} is not in the spine.'

class ReadOnly(EpubfileException):
    error_message = 'Can\'t do {} in read-only mode.'


class Epub:
    def __init__(self, epub_path, *, read_only=False):
        '''
        epub_path:
            The path to an .epub file, or to a directory that contains unzipped
             epub contents.

        read_only:
            If True, all write operations will be forbidden. The benefit is that
            the .epub file will not be extracted. This is recommended if you
            only need to read data from a book and don't need to write to it.
        '''
        epub_path = self._keep_tempdir_reference(epub_path)
        epub_path = pathclass.Path(epub_path)
        self.original_path = epub_path
        self.read_only = read_only

        if epub_path.is_dir:
            self.__init_from_dir(epub_path)
        elif self.read_only:
            self.__init_from_file_read_only(epub_path)
        else:
            self.__init_from_file(epub_path)

        opfs = self.get_opfs()
        self.opf_filepath = opfs[0]
        self.opf = self.read_opf(self.opf_filepath)

    def __init_from_dir(self, directory):
        self.is_zip = False
        self.root_directory = pathclass.Path(directory, force_sep='/')

    def __init_from_file_read_only(self, epub_path):
        # It may appear that is_zip is a synonym for read_only, but don't forget
        # that we can also open a directory in readonly mode. It's just that
        # readonly dirs don't need a special init, all they have to do is
        # forbid writes.
        self.is_zip = True
        self.root_directory = pathclass.Path(epub_path, force_sep='/')
        self.zip = zipfile.ZipFile(self.root_directory.absolute_path)

    def __init_from_file(self, epub_path):
        extract_to = tempfile.TemporaryDirectory(prefix='epubfile-')
        extract_epub(epub_path, extract_to.name)
        directory = self._keep_tempdir_reference(extract_to)
        self.__init_from_dir(directory)

    def __repr__(self):
        if self.read_only:
            return f'Epub({repr(self.root_directory.absolute_path)}, read_only=True)'
        else:
            return f'Epub({repr(self.root_directory.absolute_path)})'

    def _fopen(self, *args, **kwargs):
        '''
        Not to be confused with the high level `open_file` method, this method
        is the one that actually reads off the disk.
        '''
        if self.is_zip:
            return self._fopen_zip(*args, **kwargs)
        else:
            return self._fopen_disk(*args, **kwargs)

    def _fopen_disk(self, path, mode, *, encoding=None):
        '''
        If the book was opened as a directory, we can read files off disk with
        Python open.
        '''
        return open(path, mode, encoding=encoding)

    def _fopen_zip(self, path, mode, *, encoding=None):
        '''
        If the book was opened as a read-only zip, we can read files out of
        the zip.
        '''
        p_path = self.root_directory.spawn(path)
        if p_path in self.root_directory:
            path = p_path.relative_to(self.root_directory, simple=True)

        # Zip files always use forward slash internally, even on Windows.
        path = path.replace('\\', '/')

        if mode == 'rb':
            return self.zip.open(path, 'r')
        if mode == 'r':
            return io.TextIOWrapper(self.zip.open(path, 'r'), encoding)
        # At this time ZipFS is only used for read-only epubs anyway.
        if mode == 'wb':
            return self.zip.open(path, 'w')
        if mode == 'w':
            return io.TextIOWrapper(self.zip.open(path, 'w'), encoding)
        raise ValueError('mode should be r, w, rb, or wb.')

    def _keep_tempdir_reference(self, p):
        '''
        If the given path object is actually a tempfile.TemporaryDirectory,
        store that TD reference here so that it does not get cleaned up even
        if the caller releases it. Then return the actual filepath.
        '''
        if isinstance(p, tempfile.TemporaryDirectory):
            self._tempdir_reference = p
            p = p.name
        return p

    def assert_file_not_exists(self, filepath):
        if filepath.exists:
            existing = filepath.relative_to(self.opf_filepath.parent)
            raise FileExists(existing)

    def assert_id_not_exists(self, id):
        try:
            self.get_manifest_item(id)
            raise IDExists(id)
        except NotInManifest:
            pass

    # VALIDATION
    ############################################################################
    @writes
    def auto_correct_and_validate(self):
        # Ensure we have a mimetype file.
        mimetype_file = self.root_directory.with_child('mimetype')
        if not mimetype_file.exists:
            with self._fopen(mimetype_file.absolute_path, 'w', encoding='utf-8') as handle:
                handle.write(MIMETYPE_FILE_TEMPLATE)

        # Assert that all manifest items exist on disk.
        for item in self.get_manifest_items(soup=True):
            filepath = self.get_filepath(item['id'])
            if not filepath.exists:
                message = f'Manifest item {item["id"]} = {item["href"]} does not exist.'
                raise InvalidEpub(self.original_path, message)

    # LOADING AND SAVING
    ############################################################################
    @classmethod
    def new(cls):
        '''
        Create a new book. It will start as a temporary directory, so don't
        forget to call `save` when you are done.
        '''
        def writefile(filepath, content):
            filepath.parent.makedirs(exist_ok=True)
            # This line uses Python open instead of self._fopen because the epub
            # hasn't been instantiated yet! At this time, creating a book with
            # Epub.new always creates it as a directory. We do not support
            # creating a book directly into a fresh zip file.
            with filepath.open('w', encoding='utf-8') as handle:
                handle.write(content)

        uid = uuid.uuid4().urn

        tempdir = tempfile.TemporaryDirectory(prefix='epubfile-')
        root = pathclass.Path(tempdir.name)
        writefile(root.join('mimetype'), MIMETYPE_FILE_TEMPLATE)
        writefile(root.join('META-INF/container.xml'), CONTAINER_XML_TEMPLATE)
        writefile(root.join('OEBPS/content.opf'), OPF_TEMPLATE.format(uuid=uid))
        writefile(root.join('OEBPS/toc.ncx'), NCX_TEMPLATE.format(uuid=uid, title='Unknown', navpoints=''))
        writefile(root.join('OEBPS/Text/nav.xhtml'), NAV_XHTML_TEMPLATE.format(toc_contents=''))

        return cls(tempdir)

    @writes
    def save(self, epub_filepath):
        self.write_opf()
        self.auto_correct_and_validate()
        compress_epub(self.root_directory, epub_filepath)

    # CONTAINER & OPF
    ############################################################################
    def get_opfs(self):
        '''
        Read the container.xml to find all available OPFs (aka rootfiles).
        '''
        container = self.read_container_xml()
        rootfiles = container.find_all('rootfile')
        rootfiles = [x.get('full-path') for x in rootfiles]
        rootfiles = [self.root_directory.join(x) for x in rootfiles]
        return rootfiles

    def read_container_xml(self):
        container_xml_path = self.root_directory.join('META-INF/container.xml')
        container = self._fopen(container_xml_path.absolute_path, 'r', encoding='utf-8')
        # 'xml' and 'html.parser' seem about even here except that html.parser
        # doesn't self-close.
        container = bs4.BeautifulSoup(container, 'xml')
        return container

    def read_opf(self, rootfile):
        rootfile = pathclass.Path(rootfile, force_sep='/')
        rootfile_xml = self._fopen(rootfile.absolute_path, 'r', encoding='utf-8').read()
        # 'html.parser' preserves namespacing the best, but unfortunately it
        # botches the <meta> items because it wants them to be self-closing
        # and the string contents come out. We will fix in just a moment.
        # This is still preferable to 'xml' which handles the dc: prefixes when
        # parsing only the metadata block, but loses all namespaces when parsing
        # the whole doc. 'lxml' wraps the content in <html><body> and also
        # botches the metas so it's not any better than html.parser.
        opf = bs4.BeautifulSoup(rootfile_xml, 'html.parser')

        # Let's fix those metas.
        metas = opf.select('meta')
        for meta in metas:
            neighbor = meta.next
            if neighbor.parent != meta.parent:
                # This happens on the last meta, neighbor is outside of the manifest
                break
            if not isinstance(neighbor, bs4.element.NavigableString):
                continue
            meta.append(neighbor.extract().strip())

        return opf

    @writes
    def write_container_xml(self, container):
        if isinstance(container, bs4.BeautifulSoup):
            container = str(container)
        container_xml_path = self.root_directory.join('META-INF/container.xml')
        container_xml = self._fopen(container_xml_path.absolute_path, 'w', encoding='utf-8')
        container_xml.write(container)

    @writes
    def write_opf(self):
        with self._fopen(self.opf_filepath.absolute_path, 'w', encoding='utf-8') as rootfile:
            rootfile.write(str(self.opf))

    # FILE OPERATIONS
    ############################################################################
    @writes
    def add_file(self, id, basename, content):
        self.assert_id_not_exists(id)

        basename = os.path.basename(basename)
        mime = get_mimetype_for_basename(basename)
        directory = get_directory_for_mimetype(mime)
        directory = self.opf_filepath.parent.with_child(directory)
        directory.makedirs(exist_ok=True)
        filepath = directory.with_child(basename)

        self.assert_file_not_exists(filepath)

        if mime == 'application/xhtml+xml':
            # bs4 converts bytes to str so this must come before the handle choice.
            content = fix_xhtml(content)

        if isinstance(content, str):
            handle = self._fopen(filepath.absolute_path, 'w', encoding='utf-8')
        elif isinstance(content, bytes):
            handle = self._fopen(filepath.absolute_path, 'wb')
        else:
            raise TypeError(f'content should be str or bytes, not {type(content)}.')

        with handle:
            handle.write(content)

        href = filepath.relative_to(self.opf_filepath.parent, simple=True)
        href = urllib.parse.quote(href)

        manifest_item = make_manifest_item(id, href, mime)
        self.opf.manifest.append(manifest_item)

        if mime == 'application/xhtml+xml':
            spine_item = make_spine_item(id)
            self.opf.spine.append(spine_item)

        return id

    @writes
    def easy_add_file(self, filepath):
        '''
        Add a file from disk into the book. The manifest ID and href will be
        automatically generated.
        '''
        filepath = pathclass.Path(filepath)
        with self._fopen(filepath.absolute_path, 'rb') as handle:
            return self.add_file(
                id=filepath.basename,
                basename=filepath.basename,
                content=handle.read(),
            )

    @writes
    def delete_file(self, id):
        manifest_item = self.get_manifest_item(id)
        filepath = self.get_filepath(id)

        manifest_item.extract()
        spine_item = self.opf.spine.find('itemref', {'idref': id})
        if spine_item:
            spine_item.extract()
        os.remove(filepath.absolute_path)

    def get_filepath(self, id):
        href = self.get_manifest_item(id)['href']
        filepath = self.opf_filepath.parent.join(href)
        # TODO: In the case of a read-only zipped epub, this condition will
        # definitely fail and we won't be unquoting names that need it.
        # Double-check the consequences of this and make a patch for file
        # exists inside zip check if needed.
        if not filepath.exists:
            href = urllib.parse.unquote(href)
            filepath = self.opf_filepath.parent.join(href)
        return filepath

    def open_file(self, id, mode):
        if mode not in ('r', 'w'):
            raise ValueError(f'mode should be either r or w, not {mode}.')

        if mode == 'w' and self.read_only:
            raise ReadOnly(self.open_file.__qualname__)

        filepath = self.get_filepath(id)
        mime = self.get_manifest_item(id)['media-type']
        is_text = (
            mime in ('application/xhtml+xml', 'application/x-dtbncx+xml') or
            mime.startswith('text/')
        )

        if is_text:
            handle = self._fopen(filepath.absolute_path, mode, encoding='utf-8')
        else:
            handle = self._fopen(filepath.absolute_path, mode + 'b')

        return handle

    def read_file(self, id, *, soup=False):
        # text vs binary handled by open_file.
        content = self.open_file(id, 'r').read()
        if soup and self.get_manifest_item(id)['media-type'] == 'application/xhtml+xml':
            return fix_xhtml(content, return_soup=True)
        return content

    @writes
    def rename_file(self, id, new_basename=None, *, fix_interlinking=True):
        if isinstance(id, dict):
            basename_map = id
        else:
            if new_basename is None:
                raise TypeError('new_basename can be omitted if id is a dict.')
            basename_map = {id: new_basename}

        rename_map = {}
        for (id, new_basename) in basename_map.items():
            old_filepath = self.get_filepath(id)
            new_filepath = old_filepath.parent.with_child(new_basename)
            if not new_filepath.extension:
                new_filepath = new_filepath.add_extension(old_filepath.extension)
            self.assert_file_not_exists(new_filepath)
            os.rename(old_filepath.absolute_path, new_filepath.absolute_path)
            rename_map[old_filepath] = new_filepath

        if fix_interlinking:
            self.fix_interlinking(rename_map)
        else:
            self.fix_interlinking_opf(rename_map)

        return rename_map

    @writes
    def write_file(self, id, content):
        # text vs binary handled by open_file.
        if isinstance(content, bs4.BeautifulSoup):
            content = str(content)

        with self.open_file(id, 'w') as handle:
            handle.write(content)

    # GETTING THINGS
    ############################################################################
    def get_manifest_items(self, filter='', soup=False, spine_order=False):
        query = f'item{filter}'
        items = self.opf.manifest.select(query)

        if spine_order:
            items = {x['id']: x for x in items}
            ordered_items = []

            for spine_id in self.get_spine_order():
                ordered_items.append(items.pop(spine_id))
            ordered_items.extend(items.values())
            items = ordered_items

        if soup:
            return items

        return [x['id'] for x in items]

    def get_manifest_item(self, id):
        item = self.opf.manifest.find('item', {'id': id})
        if not item:
            raise NotInManifest(id)
        return item

    def get_fonts(self, *, soup=False):
        return self.get_manifest_items(
            filter='[media-type*="font"],[media-type*="opentype"]',
            soup=soup,
        )

    def get_images(self, *, soup=False):
        return self.get_manifest_items(
            filter='[media-type^="image/"]',
            soup=soup,
        )

    def get_media(self, *, soup=False):
        return self.get_manifest_items(
            filter='[media-type^="video/"],[media-type^="audio/"]',
            soup=soup,
        )

    def get_nav(self, *, soup=False):
        nav = self.opf.manifest.find('item', {'properties': 'nav'})
        if not nav:
            return None
        if soup:
            return nav
        return nav['id']

    def get_ncx(self, *, soup=False):
        ncx = self.opf.manifest.find('item', {'media-type': 'application/x-dtbncx+xml'})
        if not ncx:
            return None
        if soup:
            return ncx
        return ncx['id']

    def get_styles(self, *, soup=False):
        return self.get_manifest_items(
            filter='[media-type="text/css"]',
            soup=soup,
        )

    def get_texts(self, *, soup=False, skip_nav=False):
        texts = self.get_manifest_items(
            filter='[media-type="application/xhtml+xml"]',
            soup=True,
            spine_order=True,
        )
        if skip_nav:
            texts = [x for x in texts if x.get('properties') != 'nav']

        if soup:
            return texts
        return [x['id'] for x in texts]

    # COVER
    ############################################################################
    def get_cover_image(self, *, soup=False):
        cover = self.opf.manifest.find('item', {'properties': 'cover-image'})
        if cover:
            return cover if soup else cover['id']

        cover = self.opf.metadata.find('meta', {'name': 'cover'})
        if cover:
            return cover if soup else cover['content']

        return None

    @writes
    def remove_cover_image(self):
        current_cover = self.get_cover_image(soup=True)
        if not current_cover:
            return

        del current_cover['properties']

        meta = self.opf.metadata.find('meta', {'name': 'cover'})
        if meta:
            meta.extract()

    @writes
    def set_cover_image(self, id):
        if id is None:
            self.remove_cover_image()

        current_cover = self.get_cover_image(soup=True)

        if not current_cover:
            pass
        elif current_cover['id'] == id:
            return
        else:
            del current_cover['properties']

        manifest_item = self.get_manifest_item(id)
        manifest_item['properties'] = 'cover-image'

        current_meta = self.opf.metadata.find('meta', {'name': 'cover'})
        if current_meta:
            current_meta[content] = id
        else:
            meta = make_meta_item(attrs={'name': 'cover', 'content': id})
            self.opf.metadata.append(meta)

    # SPINE
    ############################################################################
    def get_spine_order(self, *, linear_only=False):
        items = self.opf.spine.find_all('itemref')
        if linear_only:
            items = [x for x in items if x.get('linear') != 'no']
        return [x['idref'] for x in items]
        return ids

    @writes
    def set_spine_order(self, ids):
        manifest_ids = self.get_manifest_items()
        # Fetch the existing entries so that we can preserve their attributes
        # while rearranging, only creating new spine entries for ids that aren't
        # already present.
        spine_items = self.opf.spine.select('itemref')
        spine_items = {item['idref']: item for item in spine_items}
        for id in ids:
            if id not in manifest_ids:
                raise NotInManifest(id)
            if id in spine_items:
                self.opf.spine.append(spine_items.pop(id))
            else:
                self.opf.spine.append(make_spine_item(id))

        # The remainder of the current spine items were not used, so pop them out.
        for spine_item in spine_items.values():
            spine_item.extract()

    def get_spine_linear(self, id):
        spine_item = self.opf.spine.find('itemref', {'idref': id})
        if not spine_item:
            raise NotInSpine(id)
        linear = spine_item.get('linear')
        linear = {None: None, 'yes': True, 'no': False}.get(linear, linear)
        return linear

    @writes
    def set_spine_linear(self, id, linear):
        '''
        Set linear to yes or no. Or pass None to remove the property.
        '''
        spine_item = self.opf.spine.find('itemref', {'idref': id})
        if not spine_item:
            raise NotInSpine(id)

        if linear is None:
            del spine_item['linear']
            return

        if isinstance(linear, str):
            if linear not in ('yes', 'no'):
                raise ValueError(f'Linear must be yes or no, not {linear}.')
        elif isinstance(linear, (bool, int)):
            linear = {True: 'yes', False: 'no'}[bool(linear)]
        else:
            raise TypeError(linear)

        spine_item['linear'] = linear

    # METADATA
    ############################################################################
    def get_authors(self):
        '''
        Thank you double_j for showing how to deal with find_all not working
        on namespaced tags.
        https://stackoverflow.com/a/44681560
        '''
        creators = self.opf.metadata.find_all({'dc:creator'})
        creators = [str(c.contents[0]) for c in creators if len(c.contents) == 1]
        return creators

    def get_dates(self):
        dates = self.opf.metadata.find_all({'dc:date'})
        dates = [str(t.contents[0]) for t in dates if len(t.contents) == 1]
        return dates

    def get_languages(self):
        languages = self.opf.metadata.find_all({'dc:language'})
        languages = [str(l.contents[0]) for l in languages if len(l.contents) == 1]
        return languages

    def get_titles(self):
        titles = self.opf.metadata.find_all({'dc:title'})
        titles = [str(t.contents[0]) for t in titles if len(t.contents) == 1]
        return titles

    @writes
    def remove_metadata_of_type(self, tag_name):
        for meta in self.opf.metadata.find_all({tag_name}):
            if meta.get('id'):
                for refines in self.opf.metadata.find_all('meta', {'refines': f'#{meta["id"]}'}):
                    refines.extract()
            meta.extract()

    @writes
    def set_languages(self, languages):
        '''
        A list like ['en', 'fr', 'ko'].
        '''
        self.remove_metadata_of_type('dc:language')
        for language in languages:
            element = f'<dc:language>{language}</dc:language>'
            element = bs4.BeautifulSoup(element, 'html.parser')
            self.opf.metadata.append(element)

    # UTILITIES
    ############################################################################
    @writes
    def fix_all_xhtml(self):
        for id in self.get_texts():
            self.write_file(id, self.read_file(id, soup=True))

    @staticmethod
    def _fix_interlinking_helper(link, rename_map, relative_to, old_relative_to=None):
        '''
        Given an old link that was found in one of the documents, and the
        rename_map, produce a new link that points to the new location.

        relative_to controls the relative pathing for the new link.
        For example, the links inside a  text document usually need to step from
        Text/ to ../Images/ to link an image. But the links inside the OPF file
        start with Images/ right away.

        old_relative_to is needed when, for example, all of the files were in a
        single directory together, and now we are splitting them into Text/,
        Images/, etc. In this case, recognizing the old link requires that we
        understand the old relative location, then we can correct it using the
        new relative location.
        '''
        if link is None:
            return None

        link = urllib.parse.urlsplit(link)
        if link.scheme:
            return None

        if old_relative_to is None:
            old_relative_to = relative_to

        new_filepath = (
            rename_map.get(link.path) or
            rename_map.get(old_relative_to.join(link.path)) or
            rename_map.get(old_relative_to.join(urllib.parse.unquote(link.path))) or
            None
        )
        if new_filepath is None:
            return None

        link = link._replace(path=new_filepath.relative_to(relative_to, simple=True))
        link = link._replace(path=urllib.parse.quote(link.path))

        return link.geturl()

    @staticmethod
    def _fix_interlinking_css_helper(tag):
        '''
        Given a <style> tag or a tag with a style="" attribute, fix interlinking
        for things like `background-image: url("");`.
        '''
        links = []
        commit = lambda: None

        if not isinstance(tag, bs4.element.Tag):
            pass

        elif tag.name == 'style' and tag.contents:
            style = tinycss2.parse_stylesheet(tag.contents[0])
            links = [
                token
                for rule in style if isinstance(rule, tinycss2.ast.QualifiedRule)
                for token in rule.content if isinstance(token, tinycss2.ast.URLToken)
            ]
            commit = lambda: tag.contents[0].replace_with(tinycss2.serialize(style))

        elif tag.get('style'):
            style = tinycss2.parse_declaration_list(tag['style'])
            links = [
                token
                for declaration in style if isinstance(declaration, tinycss2.ast.Declaration)
                for token in declaration.value if isinstance(token, tinycss2.ast.URLToken)
            ]
            commit = lambda: tag.attrs.update(style=tinycss2.serialize(style))

        return (links, commit)

    @writes
    def fix_interlinking_text(self, id, rename_map, old_relative_to=None):
        if not rename_map:
            return
        text_parent = self.get_filepath(id).parent
        soup = self.read_file(id, soup=True)
        for tag in soup.descendants:
            for link_property in HTML_LINK_PROPERTIES.get(tag.name, []):
                link = tag.get(link_property)
                link = self._fix_interlinking_helper(link, rename_map, text_parent, old_relative_to)
                if not link:
                    continue
                tag[link_property] = link

            (style_links, style_commit) = self._fix_interlinking_css_helper(tag)
            for token in style_links:
                link = token.value
                link = self._fix_interlinking_helper(link, rename_map, text_parent, old_relative_to)
                if not link:
                    continue
                token.value = link
            style_commit()

        text = str(soup)
        self.write_file(id, text)

    @writes
    def fix_interlinking_ncx(self, rename_map, old_relative_to=None):
        if not rename_map:
            return
        ncx_id = self.get_ncx()
        if not ncx_id:
            return

        ncx_parent = self.get_filepath(ncx_id).parent
        ncx = self.read_file(ncx_id)
        # 'xml' because 'lxml' and 'html.parser' lowercase the navPoint tag name.
        ncx = bs4.BeautifulSoup(ncx, 'xml')
        for point in ncx.select('navPoint > content[src]'):
            link = point['src']
            link = self._fix_interlinking_helper(link, rename_map, ncx_parent, old_relative_to)
            if not link:
                continue
            point['src'] = link

        ncx = str(ncx)
        self.write_file(ncx_id, ncx)

    @writes
    def fix_interlinking_opf(self, rename_map):
        if not rename_map:
            return
        opf_parent = self.opf_filepath.parent
        for opf_item in self.opf.select('guide > reference[href], manifest > item[href]'):
            link = opf_item['href']
            link = self._fix_interlinking_helper(link, rename_map, opf_parent)
            if not link:
                continue
            opf_item['href'] = link

    @writes
    def fix_interlinking(self, rename_map):
        if not rename_map:
            return
        self.fix_interlinking_opf(rename_map)
        for id in self.get_texts():
            self.fix_interlinking_text(id, rename_map)
        self.fix_interlinking_ncx(rename_map)

    def _set_nav_toc(self, nav_id, new_toc):
        '''
        Write the table of contents created by `generate_toc` to the nav file.
        '''
        for li in new_toc.find_all('li'):
            href = li['nav_anchor']
            atag = new_toc.new_tag('a')
            atag.append(li['text'])
            atag['href'] = href
            li.insert(0, atag)
            del li['nav_anchor']
            del li['ncx_anchor']
            del li['text']
        soup = self.read_file(nav_id, soup=True)
        toc = soup.find('nav', {'epub:type': 'toc'})
        if not toc:
            toc = soup.new_tag('nav')
            toc['epub:type'] = 'toc'
            soup.body.insert(0, toc)
        if toc.ol:
            toc.ol.extract()
        toc.append(new_toc.ol)
        self.write_file(nav_id, soup)

    def _set_ncx_toc(self, ncx_id, new_toc):
        '''
        Write the table of contents created by `generate_toc` to the ncx file.
        '''
        play_order = 1
        def li_to_navpoint(li):
            # result:
            # <navPoint id="navPoint{X}" playOrder="{X}">
            #   <navLabel>
            #     <text>{text}</text>
            #   </navLabel>
            #   <content src="{ncx_anchor}" />
            #   {children}
            # </navPoint>
            nonlocal play_order
            navpoint = new_toc.new_tag('navPoint', id=f'navPoint{play_order}', playOrder=play_order)
            play_order += 1
            label = new_toc.new_tag('navLabel')
            text = new_toc.new_tag('text')
            text.append(li['text'])
            label.append(text)
            navpoint.append(label)

            content = new_toc.new_tag('content', src=li['ncx_anchor'])
            navpoint.append(content)

            children = li.ol.children if li.ol else []
            children = [li_to_navpoint(li) for li in children]
            for child in children:
                navpoint.append(child)
            return navpoint

        # xml because we have to preserve the casing on navMap.
        soup = bs4.BeautifulSoup(self.read_file(ncx_id), 'xml')
        navmap = soup.navMap
        for child in list(navmap.children):
            child.extract()
        for li in list(new_toc.ol.children):
            navpoint = li_to_navpoint(li)
            li.insert_before(navpoint)
            li.extract()
        for navpoint in list(new_toc.ol.children):
            navmap.append(navpoint)
        self.write_file(ncx_id, soup)

    @writes
    def generate_toc(self, max_level=None, linear_only=True):
        '''
        Generate the table of contents (toc.nav and nav.xhtml) by collecting
        <h1>..<h6> throughout all of the text documents.

        max_level: If provided, only collect the headers from h1..hX, inclusive.

        linear_only: Ignore spine items that are marked as linear=no.
        '''
        def new_list(root=False):
            r = bs4.BeautifulSoup('<ol></ol>', 'html.parser')
            if root:
                return r
            return r.ol

        # Official HTML headers only go up to 6.
        if max_level is None:
            max_level = 6

        elif max_level < 1:
            raise ValueError('max_level must be >= 1.')

        header_pattern = re.compile(rf'^h[1-{max_level}]$')

        nav_id = self.get_nav()
        if nav_id:
            nav_filepath = self.get_filepath(nav_id)

        ncx_id = self.get_ncx()
        if ncx_id:
            ncx_filepath = self.get_filepath(ncx_id)

        if not nav_id and not ncx_id:
            return

        # Note: The toc generated by the upcoming loop is in a sort of agnostic
        # format, since it needs to be converted into nav.html and toc.ncx which
        # have different structural requirements. The attributes that I'm using
        # in this initial toc object DO NOT represent any part of the epub format.
        toc = new_list(root=True)

        current_list = toc.ol
        current_list['level'] = None

        spine = self.get_spine_order(linear_only=linear_only)
        spine = [s for s in spine if s != nav_id]

        for file_id in spine:
            file_path = self.get_filepath(file_id)
            soup = self.read_file(file_id, soup=True)

            headers = soup.find_all(header_pattern)
            for (toc_line_index, header) in enumerate(headers, start=1):
                # 'hX' -> X
                level = int(header.name[1])

                header['id'] = f'toc_{toc_line_index}'

                toc_line = toc.new_tag('li')
                toc_line['text'] = header.text

                # In Lithium, the TOC drawer only remembers your position if
                # the page that you're reading corresponds to a TOC entry
                # exactly. Which is to say, if you left off on page5.html,
                # there needs to be a TOC line with href="page5.html" or else
                # the TOC drawer will be in the default position at the top of
                # the list and not highlight the current chapter. Any #anchor
                # in the href will break this feature. So, this code will make
                # the first <hX> on a given page not have an #anchor. If you
                # have a significant amount of text on the page before this
                # header, then this will look bad. But for the majority of
                # cases I expect the first header on the page will be at the
                # very top, or near enough that the Lithium fix is still
                # worthwhile.
                if toc_line_index == 1:
                    hash_anchor = ''
                else:
                    hash_anchor = f'#{header["id"]}'

                if nav_id:
                    relative = file_path.relative_to(nav_filepath.parent, simple=True)
                    toc_line['nav_anchor'] = f'{relative}{hash_anchor}'
                if ncx_id:
                    relative = file_path.relative_to(ncx_filepath.parent, simple=True)
                    toc_line['ncx_anchor'] = f'{relative}{hash_anchor}'

                if current_list['level'] is None:
                    current_list['level'] = level

                while level < current_list['level']:
                    # Because the sub-<ol> are actually a child of the last
                    # <li> of the previous <ol>, we must .parent twice.
                    # The second .parent is conditional because if the current
                    # list is toc.ol, then parent is a Soup document object, and
                    # parenting again would be a mistake. We'll recover from
                    # this in just a moment.
                    current_list = current_list.parent
                    if current_list.name == 'li':
                        current_list = current_list.parent
                    # If the file has headers in a non-ascending order, like the
                    # first header is an h4 and then an h1 comes later, then
                    # this while loop would keep attempting to climb the .parent
                    # which would take us too far, off the top of the tree.
                    # So, if we reach `current_list == toc.ol` then we've
                    # reached the root and should stop climbing. At that point
                    # we can just snap current_level and use the root list again.
                    # In the resulting toc, that initial h4 would have the same
                    # toc depth as the later h1 since it never had parents.
                    if current_list == toc:
                        current_list['level'] = level
                        current_list = toc.ol

                if level > current_list['level']:
                    # In order to properly render nested <ol>, you're supposed
                    # to make the new <ol> a child of the last <li> of the
                    # previous <ol>. NOT a child of the prev <ol> directly.
                    # Don't worry, .children can never be empty because on the
                    # first <li> this condition can never occur, and new <ol>s
                    # always receive a child right after being created.
                    _l = new_list()
                    _l['level'] = level
                    final_li = list(current_list.children)[-1]
                    final_li.append(_l)
                    current_list = _l

                current_list.append(toc_line)

            # We have to save the id="toc_X" that we gave to all the headers.
            self.write_file(file_id, soup)

        for ol in toc.find_all('ol'):
            del ol['level']

        if nav_id:
            self._set_nav_toc(nav_id, copy.copy(toc))

        if ncx_id:
            self._set_ncx_toc(ncx_id, copy.copy(toc))

    @writes
    def move_nav_to_end(self):
        '''
        Move the nav.xhtml file to the end and set its linear=no.
        '''
        nav = self.get_nav()
        if not nav:
            return

        spine = self.get_spine_order()

        try:
            index = spine.index(nav)
            spine.append(spine.pop(index))
        except ValueError:
            spine.append(nav)

        self.set_spine_order(spine)
        self.set_spine_linear(nav, False)

    @writes
    def normalize_directory_structure(self):
        # This must come before the opf rewrite because that would affect the
        # location of all all manifest item hrefs.
        manifest_items = self.get_manifest_items(soup=True)
        old_filepaths = {item['id']: self.get_filepath(item['id']) for item in manifest_items}
        old_ncx = self.get_ncx()
        try:
            old_ncx_parent = self.get_filepath(self.get_ncx()).parent
        except Exception:
            old_ncx_parent = None

        if self.opf_filepath.parent == self.root_directory:
            oebps = self.root_directory.with_child('OEBPS')
            oebps.makedirs(exist_ok=True)
            self.write_opf()
            new_opf_path = oebps.with_child(self.opf_filepath.basename)
            os.rename(self.opf_filepath.absolute_path, new_opf_path.absolute_path)
            container = self.read_container_xml()
            rootfile = container.find('rootfile', {'full-path': self.opf_filepath.basename})
            rootfile['full-path'] = new_opf_path.relative_to(self.root_directory, simple=True)
            self.write_container_xml(container)
            self.opf_filepath = new_opf_path

        rename_map = {}
        for manifest_item in manifest_items:
            old_filepath = old_filepaths[manifest_item['id']]

            directory = get_directory_for_mimetype(manifest_item['media-type'])
            directory = self.opf_filepath.parent.with_child(directory)
            if directory.exists:
                # On Windows, this will fix any incorrect casing.
                # On Linux it is inert.
                os.rename(directory.absolute_path, directory.absolute_path)
            else:
                directory.makedirs()

            new_filepath = directory.with_child(old_filepath.basename)
            if new_filepath.absolute_path != old_filepath.absolute_path:
                rename_map[old_filepath] = new_filepath
                os.rename(old_filepath.absolute_path, new_filepath.absolute_path)
            manifest_item['href'] = new_filepath.relative_to(self.opf_filepath.parent, simple=True)

        self.fix_interlinking_opf(rename_map)
        for id in self.get_texts():
            self.fix_interlinking_text(id, rename_map, old_relative_to=old_filepaths[id].parent)
        self.fix_interlinking_ncx(rename_map, old_relative_to=old_ncx_parent)

    @writes
    def normalize_opf(self):
        for tag in self.opf.descendants:
            if tag.name:
                tag.name = tag.name.replace('opf:', '')
        for item in self.get_manifest_items(soup=True):
            if item['href'] in ['toc.ncx', 'Misc/toc.ncx']:
                item['media-type'] = 'application/x-dtbncx+xml'


# COMMAND LINE TOOLS
################################################################################
import argparse
import html
import random
import string
import sys

from voussoirkit import betterhelp
from voussoirkit import winglob

DOCSTRING = '''
Epubfile
The simple python .epub scripting tool.

{addfile}

{covercomesfirst}

{exec}

{generate_toc}

{holdit}

{merge}

{normalize}

{setfont}

TO SEE DETAILS ON EACH COMMAND, RUN
> epubfile.py <command>
'''

SUB_DOCSTRINGS = dict(
addfile='''
addfile:
    Add files into the book.

    > epubfile.py addfile book.epub page1.html image.jpg
'''.strip(),

covercomesfirst='''
covercomesfirst:
    Rename the cover image file so that it is the alphabetically-first image.

    > epubfile.py covercomesfirst book.epub

    I use CBXShell to get thumbnails of epub files on Windows, and because it
    is generalized for zip files and doesn't read epub metadata, alphabetized
    mode works best for getting epub covers as icons.

    In my testing, CBXShell considers the image's whole path and not just the
    basename, so you may want to consider normalizing the directory structure
    first, otherwise some /a/image.jpg will always be before /images/cover.jpg.
'''.strip(),

exec='''
exec:
    Execute a snippet of Python code against the book.

    > epubfile.py exec book.epub --command "book._____()"
'''.strip(),

generate_toc='''
generate_toc:
    Regenerate the toc.ncx and nav.xhtml based on html <hX> headers in the text.

    > epubfile.py generate_toc book.epub <flags>

    flags:
    --max_level X:
        Only generate toc entries for headers up to level X.
        That is, h1, h2, ... hX.
'''.strip(),

holdit='''
holdit:
    Extract the book so that you can manually edit the files on disk, then save
    the changes back into the original file.

    > epubfile.py holdit book.epub
'''.strip(),

merge='''
merge:
    Merge multiple books into one.

    > epubfile.py merge book1.epub book2.epub --output final.epub <flags>

    flags:
    --demote_headers:
        All h1 in the book will be demoted to h2, and so forth. So that the
        headerfiles are the only h1s and the table of contents will generate
        with a good hierarchy.

    --headerfile:
        Add a file before each book with an <h1> containing its title.

    --number_headerfile:
        In the headerfile, the <h1> will start with the book's index, like
        "01. First Book"

    -y | --autoyes:
        Overwrite the output file without prompting.
'''.strip(),

normalize='''
normalize:
    Rename files and directories in the book to match a common structure.

    Moves all book content from / into /OEBPS and sorts files into
    subdirectories by type: Text, Images, Styles, etc.

    > epubfile.py normalize book.epub
'''.strip(),

setfont='''
setfont:
    Set the font for every page in the whole book.

    A stylesheet called epubfile_setfont.css will be created that sets
    * { font-family: ... !important } with a font file of your choice.

    > epubfile.py setfont book.epub font.ttf
'''.strip(),
)

DOCSTRING = betterhelp.add_previews(DOCSTRING, SUB_DOCSTRINGS)

def random_string(length, characters=string.ascii_lowercase):
    return ''.join(random.choice(characters) for x in range(length))

def addfile_argparse(args):
    book = Epub(args.epub)

    for pattern in args.files:
        for file in winglob.glob(pattern):
            print(f'Adding file {file}.')
            file = pathclass.Path(file)
            try:
                book.easy_add_file(file)
            except (IDExists, FileExists) as exc:
                rand_suffix = random_string(3, string.digits)
                base = file.replace_extension('').basename
                id = f'{base}_{rand_suffix}'
                basename = f'{base}_{rand_suffix}{file.extension.with_dot}'
                content = file.open('rb').read()
                book.add_file(id, basename, content)

    book.move_nav_to_end()
    book.save(args.epub)

def covercomesfirst(book):
    basenames = {i: book.get_filepath(i).basename for i in book.get_images()}
    if len(basenames) <= 1:
        return

    cover_image = book.get_cover_image()
    if not cover_image:
        return

    cover_basename = book.get_filepath(cover_image).basename

    cover_index = sorted(basenames.values()).index(cover_basename)
    if cover_index == 0:
        return

    rename_map = basenames.copy()

    if not cover_basename.startswith('!'):
        cover_basename = '!' + cover_basename
        rename_map[cover_image] = cover_basename
    else:
        rename_map.pop(cover_image)

    for (id, basename) in rename_map.copy().items():
        if id == cover_image:
            continue
        if basename > cover_basename:
            rename_map.pop(id)
            continue
        if basename < cover_basename and basename.startswith('!'):
            basename = basename.lstrip('!')
            rename_map[id] = basename
        if basename < cover_basename or basename.startswith('.'):
            basename = '_' + basename
            rename_map[id] = basename

    book.rename_file(rename_map)

def covercomesfirst_argparse(args):
    epubs = [epub for pattern in args.epubs for epub in winglob.glob(pattern)]
    for epub in epubs:
        print(epub)
        book = Epub(epub)
        covercomesfirst(book)
        book.save(args.epub)

def exec_argparse(args):
    epubs = [epub for pattern in args.epubs for epub in winglob.glob(pattern)]
    for epub in epubs:
        print(epub)
        book = Epub(epub)
        exec(args.command)
        book.save(epub)

def generate_toc_argparse(args):
    epubs = [epub for pattern in args.epubs for epub in winglob.glob(pattern)]
    books = []
    for epub in epubs:
        book = Epub(epub)
        book.generate_toc(max_level=int(args.max_level) if args.max_level else None)
        book.save(epub)

def holdit_argparse(args):
    epubs = [epub for pattern in args.epubs for epub in winglob.glob(pattern)]
    books = []
    for epub in epubs:
        book = Epub(epub)
        print(f'{epub} = {book.root_directory.absolute_path}')
        books.append((epub, book))

    input('Press Enter when ready.')
    for (epub, book) in books:
        # Saving re-writes the opf from memory, which might undo any manual changes.
        # So let's re-read it first.
        book.read_opf(book.opf_filepath)
        book.save(epub)

def merge(
        input_filepaths,
        output_filename,
        demote_headers=False,
        do_headerfile=False,
        number_headerfile=False,
    ):
    book = Epub.new()

    input_filepaths = [pathclass.Path(p) for pattern in input_filepaths for p in winglob.glob(pattern)]
    index_length = len(str(len(input_filepaths)))
    rand_prefix = random_string(3, string.digits)

    # Number books from 1 for human sanity.
    for (index, input_filepath) in enumerate(input_filepaths, start=1):
        print(f'Merging {input_filepath.absolute_path}.')
        prefix = f'{rand_prefix}_{index:>0{index_length}}_{{}}'
        input_book = Epub(input_filepath)
        input_book.normalize_directory_structure()

        input_ncx = input_book.get_ncx()
        input_nav = input_book.get_nav()
        manifest_ids = input_book.get_manifest_items(spine_order=True)
        manifest_ids = [x for x in manifest_ids if x not in (input_ncx, input_nav)]

        basename_map = {}
        for id in manifest_ids:
            old_basename = input_book.get_filepath(id).basename
            new_basename = prefix.format(old_basename)
            basename_map[id] = new_basename

        # Don't worry, we're not going to save over the input book!
        input_book.rename_file(basename_map)

        if do_headerfile:
            content = ''
            try:
                title = input_book.get_titles()[0]
            except IndexError:
                title = input_filepath.replace_extension('').basename

            try:
                year = input_book.get_dates()[0]
            except IndexError:
                pass
            else:
                title = f'{title} ({year})'

            if number_headerfile:
                title = f'{index:>0{index_length}}. {title}'

            content += f'<h1>{html.escape(title)}</h1>'

            try:
                author = input_book.get_authors()[0]
                content += f'<p>{html.escape(author)}</p>'
            except IndexError:
                pass

            headerfile_id = prefix.format('headerfile')
            headerfile_basename = prefix.format('headerfile.html')
            book.add_file(headerfile_id, headerfile_basename, content)

        for id in manifest_ids:
            new_id = prefix.format(id)
            new_basename = basename_map[id]
            if demote_headers:
                content = input_book.read_file(id, soup=True)
                if isinstance(content, bs4.BeautifulSoup):
                    content = demote_xhtml_headers(content)
            else:
                content = input_book.read_file(id)
            book.add_file(new_id, new_basename, content)

    book.move_nav_to_end()
    book.save(output_filename)

def merge_argparse(args):
    if os.path.exists(args.output):
        if not (args.autoyes or getpermission.getpermission(f'Overwrite {args.output}?')):
            raise ValueError(f'{args.output} exists.')

    return merge(
        input_filepaths=args.epubs,
        output_filename=args.output,
        demote_headers=args.demote_headers,
        do_headerfile=args.headerfile,
        number_headerfile=args.number_headerfile,
    )

def normalize_argparse(args):
    epubs = [epub for pattern in args.epubs for epub in winglob.glob(pattern)]
    for epub in epubs:
        print(epub)
        book = Epub(epub)
        book.normalize_opf()
        book.normalize_directory_structure()
        book.move_nav_to_end()
        book.save(epub)

def setfont_argparse(args):
    book = Epub(args.epub)

    css_id = 'epubfile_setfont'
    css_basename = 'epubfile_setfont.css'

    try:
        book.assert_id_not_exists(css_id)
    except IDExists:
        if not getpermission.getpermission(f'Overwrite {css_id}?'):
            return
        book.delete_file(css_id)

    font = pathclass.Path(args.font)

    for existing_font in book.get_fonts():
        font_path = book.get_filepath(existing_font)
        if font_path.basename == font.basename:
            font_id = existing_font
            break
    else:
        font_id = book.easy_add_file(font)
        font_path = book.get_filepath(font_id)

    # The font_path may have come from an existing font in the book, so we have
    # no guarantees about its path layout. The css file, however, is definitely
    # going to be inside OEBPS/Styles since we're the ones creating it.
    # So, we should be getting the correct number of .. in the relative path.
    family = font_path.basename
    relative = font_path.relative_to(book.opf_filepath.parent.with_child('Styles'))

    css = f'''
    @font-face {{
    font-family: '{family}';
    font-weight: normal;
    font-style: normal;
    src: url("{relative}");
    }}

    * {{
        font-family: '{family}' !important;
    }}
    '''

    book.add_file(
        id=css_id,
        basename=css_basename,
        content=css,
    )
    css_path = book.get_filepath(css_id)

    for text_id in book.get_texts():
        text_path = book.get_filepath(text_id)
        relative = css_path.relative_to(text_path)
        soup = book.read_file(text_id, soup=True)
        head = soup.head
        if head.find('link', {'id': css_id}):
            continue
        link = soup.new_tag('link')
        link['id'] = css_id
        link['href'] = css_path.relative_to(text_path.parent)
        link['rel'] = 'stylesheet'
        link['type'] = 'text/css'
        head.append(link)
        book.write_file(text_id, soup)

    book.save(args.epub)

def main(argv):
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers()

    p_addfile = subparsers.add_parser('addfile')
    p_addfile.add_argument('epub')
    p_addfile.add_argument('files', nargs='+', default=[])
    p_addfile.set_defaults(func=addfile_argparse)

    p_covercomesfirst = subparsers.add_parser('covercomesfirst')
    p_covercomesfirst.add_argument('epubs', nargs='+', default=[])
    p_covercomesfirst.set_defaults(func=covercomesfirst_argparse)

    p_exec = subparsers.add_parser('exec')
    p_exec.add_argument('epubs', nargs='+', default=[])
    p_exec.add_argument('--command', dest='command', default=None, required=True)
    p_exec.set_defaults(func=exec_argparse)

    p_generate_toc = subparsers.add_parser('generate_toc')
    p_generate_toc.add_argument('epubs', nargs='+', default=[])
    p_generate_toc.add_argument('--max_level', '--max-level', dest='max_level', default=None)
    p_generate_toc.set_defaults(func=generate_toc_argparse)

    p_holdit = subparsers.add_parser('holdit')
    p_holdit.add_argument('epubs', nargs='+', default=[])
    p_holdit.set_defaults(func=holdit_argparse)

    p_merge = subparsers.add_parser('merge')
    p_merge.add_argument('epubs', nargs='+', default=[])
    p_merge.add_argument('--output', dest='output', default=None, required=True)
    p_merge.add_argument('--headerfile', dest='headerfile', action='store_true')
    p_merge.add_argument('--demote_headers', '--demote-headers', dest='demote_headers', action='store_true')
    p_merge.add_argument('--number_headerfile', '--number-headerfile', dest='number_headerfile', action='store_true')
    p_merge.add_argument('-y', '--autoyes', dest='autoyes', action='store_true')
    p_merge.set_defaults(func=merge_argparse)

    p_normalize = subparsers.add_parser('normalize')
    p_normalize.add_argument('epubs', nargs='+', default=[])
    p_normalize.set_defaults(func=normalize_argparse)

    p_setfont = subparsers.add_parser('setfont')
    p_setfont.add_argument('epub')
    p_setfont.add_argument('font')
    p_setfont.set_defaults(func=setfont_argparse)

    return betterhelp.subparser_main(
        argv,
        parser,
        main_docstring=DOCSTRING,
        sub_docstrings=SUB_DOCSTRINGS,
    )

if __name__ == '__main__':
    raise SystemExit(main(sys.argv[1:]))
