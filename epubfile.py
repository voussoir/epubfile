import mimetypes
import os
import re
import tempfile
import urllib.parse
import uuid
import zipfile

import bs4
import tinycss2

from voussoirkit import pathclass

MIMETYPE_CONTENT = 'application/epub+zip'

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
    'html': 'application/xhtml+xml',
    'xhtml': 'application/xhtml+xml',
    'htm': 'application/xhtml+xml',
    'smi': 'application/smil+xml',
    'smil': 'application/smil+xml',
    'sml': 'application/smil+xml',
    'pls': 'application/pls+xml',
    'otf': 'font/otf',
    'ttf': 'font/ttf',
    'woff': 'font/woff',
    'woff2': 'font/woff2',
}

MIMETYPE_DIRECTORIES = {
    'application/x-dtbncx+xml': '.',
    'application/font-sfnt': 'Fonts',
    'application/xhtml+xml': 'Text',
    'font': 'Fonts',
    'image': 'Images',
    'text/css': 'Styles',
    'audio': 'Audio',
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
   <text>Unknown</text>
</docTitle>
<navMap>
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
    <ol>
    </ol>
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
    meta_item = bs4.BeautifulSoup(meta_item, 'html.parser')
    if attrs:
        meta_item.attrs.update(attrs)
    return meta_item.meta

def make_spine_item(id):
    spine_item = f'<itemref idref="{id}"/>'
    # 'html.parser' just for having the simplest output.
    spine_item = bs4.BeautifulSoup(spine_item, 'html.parser')
    return spine_item.itemref


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

class FileExists(EpubfileException):
    error_message = 'There is already a file at {}.'

class IDExists(EpubfileException):
    error_message = 'There is already a file with id {}.'

class NotInManifest(EpubfileException):
    error_message = '{} is not in the manifest.'

class NotInSpine(EpubfileException):
    error_message = '{} is not in the spine.'


class Epub:
    def __init__(self, directory):
        if isinstance(directory, tempfile.TemporaryDirectory):
            self._tempdir_reference = directory
            directory = directory.name

        self.root_directory = pathclass.Path(directory, force_sep='/')

        self.opf_filepath = None
        self.opf = None

        self.read_opf(self.get_opfs()[0])

    def __repr__(self):
        return f'Epub({repr(self.root_directory.absolute_path)})'

    def assert_file_not_exists(self, filepath):
        if filepath.exists:
            existing = filepath.relative_to(self.opf_filepath.parent)
            raise FileExists(existing)

    def assert_id_not_exists(self, id):
        if self.opf.manifest.find('item', {'id': id}):
            raise IDExists(id)

    # LOADING AND SAVING
    ############################################################################
    @classmethod
    def new(cls):
        def writefile(filepath, content):
            os.makedirs(filepath.parent.absolute_path, exist_ok=True)
            with open(filepath.absolute_path, 'w', encoding='utf-8') as handle:
                handle.write(content)

        uid = uuid.uuid4().urn

        tempdir = tempfile.TemporaryDirectory(prefix='epubfile-')
        root = pathclass.Path(tempdir.name)
        writefile(root.join('mimetype'), MIMETYPE_FILE_TEMPLATE)
        writefile(root.join('META-INF/container.xml'), CONTAINER_XML_TEMPLATE)
        writefile(root.join('OEBPS/content.opf'), OPF_TEMPLATE.format(uuid=uid))
        writefile(root.join('OEBPS/toc.ncx'), NCX_TEMPLATE.format(uuid=uid))
        writefile(root.join('OEBPS/Text/nav.xhtml'), NAV_XHTML_TEMPLATE)

        return cls(tempdir)

    @classmethod
    def open(cls, epub_filepath):
        extract_to = tempfile.TemporaryDirectory(prefix='epubfile-')
        extract_epub(epub_filepath, extract_to.name)
        return cls(extract_to)

    def save(self, epub_filepath):
        self.write_opf()
        compress_epub(self.root_directory, epub_filepath)

    # CONTAINER & OPF
    ############################################################################
    def get_opfs(self):
        container = self.read_container_xml()
        rootfiles = container.find_all('rootfile')
        rootfiles = [x.get('full-path') for x in rootfiles]
        rootfiles = [self.root_directory.join(x) for x in rootfiles]
        return rootfiles

    def read_container_xml(self):
        container_xml_path = self.root_directory.join('META-INF/container.xml')
        container = open(container_xml_path.absolute_path, 'r', encoding='utf-8')
        # 'xml' and 'html.parser' seem about even here except that html.parser doesn't self-close.
        container = bs4.BeautifulSoup(container, 'xml')
        return container

    def read_opf(self, rootfile):
        rootfile = pathclass.Path(rootfile, force_sep='/')
        rootfile_xml = open(rootfile.absolute_path, 'r', encoding='utf-8').read()
        # 'html.parser' preserves namespacing the best, but unfortunately it
        # botches the <meta> items because it wants them to be self-closing
        # and the string contents come out. We will fix in just a moment.
        # This is still preferable to 'xml' which handles the dc: prefixes when
        # parsing only the metadata block, but loses all namespaces when parsing
        # the whole doc. 'lxml' wraps the content in <html><body> and also
        # botches the metas so it's not any better than html.parser.
        self.opf = bs4.BeautifulSoup(rootfile_xml, 'html.parser')
        # Let's fix those metas.
        metas = self.opf.select('meta')
        for meta in metas:
            neighbor = meta.next
            if neighbor.parent != meta.parent:
                break
            if not isinstance(neighbor, bs4.element.NavigableString):
                continue
            meta.append(neighbor.extract().strip())

        self.opf_filepath = rootfile
        return self.opf

    def write_container_xml(self, container):
        if isinstance(container, bs4.BeautifulSoup):
            container = str(container)
        container_xml_path = self.root_directory.join('META-INF/container.xml')
        container_xml = open(container_xml_path.absolute_path, 'w', encoding='utf-8')
        container_xml.write(container)

    def write_opf(self):
        with open(self.opf_filepath.absolute_path, 'w', encoding='utf-8') as rootfile:
            rootfile.write(str(self.opf))

    # FILE OPERATIONS
    ############################################################################
    def add_file(self, id, basename, content):
        self.assert_id_not_exists(id)

        basename = os.path.basename(basename)
        mime = get_mimetype_for_basename(basename)
        directory = get_directory_for_mimetype(mime)
        directory = self.opf_filepath.parent.with_child(directory)
        os.makedirs(directory.absolute_path, exist_ok=True)
        filepath = directory.with_child(basename)

        self.assert_file_not_exists(filepath)

        if mime == 'application/xhtml+xml':
            # bs4 converts bytes to str so this must come before the handle choice.
            content = fix_xhtml(content)

        if isinstance(content, str):
            handle = open(filepath.absolute_path, 'w', encoding='utf-8')
        elif isinstance(content, bytes):
            handle = open(filepath.absolute_path, 'wb')
        else:
            raise TypeError(type(content))

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

    def easy_add_file(self, filepath):
        filepath = pathclass.Path(filepath)
        with open(filepath.absolute_path, 'rb') as handle:
            self.add_file(
                id=filepath.basename,
                basename=filepath.basename,
                content=handle.read(),
            )

    def delete_file(self, id):
        os.remove(self.get_filepath(id).absolute_path)
        spine_item = self.opf.spine.find('itemref', {'idref': id})
        if spine_item:
            spine_item.extract()
        manifest_item = self.opf.manifest.find('item', {'id': id})
        manifest_item.extract()

    def get_filepath(self, id):
        href = self.opf.manifest.find('item', {'id': id})['href']
        filepath = self.opf_filepath.parent.join(href)
        if not filepath.exists:
            href = urllib.parse.unquote(href)
            filepath = self.opf_filepath.parent.join(href)
        return filepath

    def open_file(self, id, mode):
        if mode not in ('r', 'w'):
            raise ValueError(f'Mode {mode} should be either r or w.')

        filepath = self.get_filepath(id)
        mime = self.opf.manifest.find('item', {'id': id})['media-type']
        is_text = (
            mime in ('application/xhtml+xml', 'application/x-dtbncx+xml') or
            mime.startswith('text/')
        )

        if is_text:
            handle = open(filepath.absolute_path, mode, encoding='utf-8')
        else:
            handle = open(filepath.absolute_path, mode + 'b')

        return handle

    def read_file(self, id, *, soup=False):
        # text vs binary handled by open_file.
        content = self.open_file(id, 'r').read()
        if soup and self.get_manifest_item(id)['media-type'] == 'application/xhtml+xml':
            return fix_xhtml(content, return_soup=True)
        return content

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
            filter='[media-type^="application/font"],[media-type^="font/"]',
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

    def remove_cover_image(self):
        current_cover = self.get_cover_image(soup=True)
        if not current_cover:
            return

        del current_cover['properties']

        meta = self.opf.metadata.find('meta', {'name': 'cover'})
        if meta:
            meta.extract()

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
    def get_spine_order(self, *, only_linear=False):
        items = self.opf.spine.find_all('itemref')
        if only_linear:
            items = [x for x in items if x.get('linear') != 'no']
        return [x['idref'] for x in items]
        return ids

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

    def get_languages(self):
        languages = self.opf.metadata.find_all({'dc:language'})
        languages = [str(l.contents[0]) for l in languages if len(l.contents) == 1]
        return languages

    def get_titles(self):
        titles = self.opf.metadata.find_all({'dc:title'})
        titles = [str(t.contents[0]) for t in titles if len(t.contents) == 1]
        return titles

    # UTILITIES
    ############################################################################
    def fix_all_xhtml(self):
        for id in self.get_texts():
            self.write_file(id, self.read_file(id, soup=True))

    @staticmethod
    def _fix_interlinking_helper(link, rename_map, relative_to, old_relative_to=None):
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

    def fix_interlinking_text(self, id, rename_map, old_relative_to=None):
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

    def fix_interlinking_ncx(self, rename_map, old_relative_to=None):
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

    def fix_interlinking_opf(self, rename_map):
        opf_parent = self.opf_filepath.parent
        for opf_item in self.opf.select('guide > reference[href], manifest > item[href]'):
            link = opf_item['href']
            link = self._fix_interlinking_helper(link, rename_map, opf_parent)
            if not link:
                continue
            opf_item['href'] = link

    def fix_interlinking(self, rename_map):
        self.fix_interlinking_opf(rename_map)
        for id in self.get_texts():
            self.fix_interlinking_text(id, rename_map)
        self.fix_interlinking_ncx(rename_map)

    def move_nav_to_end(self):
        '''
        Move the nav.xhtml file to the end and set linear=no.
        '''
        nav = self.get_nav()
        if not nav:
            return

        spine = self.get_spine_order()
        for (index, id) in enumerate(spine):
            if id == nav:
                spine.append(spine.pop(index))
                break
        self.set_spine_order(spine)

        self.set_spine_linear(nav, False)

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
            os.makedirs(oebps.absolute_path, exist_ok=True)
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
            os.makedirs(directory.absolute_path, exist_ok=True)

            new_filepath = directory.with_child(old_filepath.basename)
            rename_map[old_filepath] = new_filepath
            os.rename(old_filepath.absolute_path, new_filepath.absolute_path)
            manifest_item['href'] = new_filepath.relative_to(self.opf_filepath.parent, simple=True)

        self.fix_interlinking_opf(rename_map)
        for id in self.get_texts():
            self.fix_interlinking_text(id, rename_map, old_relative_to=old_filepaths[id].parent)
        self.fix_interlinking_ncx(rename_map, old_relative_to=old_ncx_parent)


# COMMAND LINE TOOLS
################################################################################
import argparse
import html
import random
import string
import sys

from voussoirkit import betterhelp

DOCSTRING = '''
{addfile}

{covercomesfirst}

{holdit}

{merge}

{normalize}
'''.lstrip()

SUB_DOCSTRINGS = {
'addfile':
'''
addfile:
    Add files into the book.

    > epubfile.py addfile book.epub page1.html image.jpg
'''.strip(),

'covercomesfirst':
'''
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

'holdit':
'''
holdit:
    Extract the book and leave it open for manual editing, then save.

    > epubfile.py holdit book.epub
''',

'merge':
'''
merge:
    Merge multiple books into one.

    > epubfile.py merge book1.epub book2.epub --output final.epub <flags>

    flags:
    --headerfile:
        Add a file before each book with an <h1> containing its title.

    -y | --autoyes:
        Overwrite the output file without prompting.
'''.strip(),

'normalize':
'''
normalize:
    Rename files and directories in the book to match a common structure.

    Moves all book content from / into /OEBPS and sorts files into
    subdirectories by type: Text, Images, Styles, etc.

    > epubfile.py normalize book.epub
'''.strip()
}

DOCSTRING = betterhelp.add_previews(DOCSTRING, SUB_DOCSTRINGS)

def random_string(length, characters=string.ascii_lowercase):
    return ''.join(random.choice(characters) for x in range(length))

def addfile_argparse(args):
    book = Epub.open(args.epub)

    for file in args.files:
        print(f'Adding file {file}.')
        file = pathclass.Path(file)
        try:
            book.easy_add_file(file)
        except (IDExists, FileExists) as exc:
            rand_suffix = random_string(3, string.digits)
            base = file.replace_extension('').basename
            id = f'{base}_{rand_suffix}'
            basename = f'{base}_{rand_suffix}{file.dot_extension}'
            content = open(file.absolute_path, 'rb').read()
            book.add_file(id, basename, content)

    book.move_nav_to_end()
    book.save(args.epub)

def covercomesfirst_argparse(args):
    book = Epub.open(args.epub)
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

    book.save(args.epub)

def holdit_argparse(args):
    book = Epub.open(args.epub)
    print(book.root_directory.absolute_path)
    input('Press Enter when ready.')
    book.save(args.epub)

def merge(input_filepaths, output_filename, do_headerfile=False):
    book = Epub.new()

    index_length = len(str(len(input_filepaths)))
    rand_prefix = random_string(3, string.digits)

    input_filepaths = [pathclass.Path(p) for p in input_filepaths]

    for (index, input_filepath) in enumerate(input_filepaths):
        print(f'Merging {input_filepath}.')
        prefix = f'{rand_prefix}_{index:>0{index_length}}_{{}}'
        input_book = Epub.open(input_filepath)
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

        # Don't worry, we're not going to save this!
        input_book.rename_file(basename_map)

        if do_headerfile:
            content = ''
            try:
                title = input_book.get_titles()[0]
            except IndexError:
                title = input_filepath.replace_extension('').basename
            finally:
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
            new_id = f'{rand_prefix}_{index:>0{index_length}}_{id}'
            new_basename = basename_map[id]
            book.add_file(new_id, new_basename, input_book.read_file(id))

    book.move_nav_to_end()
    book.save(output_filename)

def merge_argparse(args):
    if os.path.exists(args.output):
        ok = args.autoyes
        if not ok:
            ok = input(f'Overwrite {args.output}? y/n\n>').lower() in ('y', 'yes')
        if not ok:
            raise ValueError(f'{args.output} exists.')

    return merge(input_filepaths=args.epubs, output_filename=args.output, do_headerfile=args.headerfile)

def normalize_argparse(args):
    book = Epub.open(args.epub)
    book.normalize_directory_structure()
    book.save(args.epub)

@betterhelp.subparser_betterhelp(main_docstring=DOCSTRING, sub_docstrings=SUB_DOCSTRINGS)
def main(argv):
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers()

    p_addfile = subparsers.add_parser('addfile')
    p_addfile.add_argument('epub')
    p_addfile.add_argument('files', nargs='+', default=[])
    p_addfile.set_defaults(func=addfile_argparse)

    p_covercomesfirst = subparsers.add_parser('covercomesfirst')
    p_covercomesfirst.add_argument('epub')
    p_covercomesfirst.set_defaults(func=covercomesfirst_argparse)

    p_holdit = subparsers.add_parser('holdit')
    p_holdit.add_argument('epub')
    p_holdit.set_defaults(func=holdit_argparse)

    p_merge = subparsers.add_parser('merge')
    p_merge.add_argument('epubs', nargs='+', default=[])
    p_merge.add_argument('--output', dest='output', default=None, required=True)
    p_merge.add_argument('--headerfile', dest='headerfile', action='store_true')
    p_merge.add_argument('-y', '--autoyes', dest='autoyes', action='store_true')
    p_merge.set_defaults(func=merge_argparse)

    p_normalize = subparsers.add_parser('normalize')
    p_normalize.add_argument('epub')
    p_normalize.set_defaults(func=normalize_argparse)

    args = parser.parse_args(argv)
    args.func(args)

if __name__ == '__main__':
    raise SystemExit(main(sys.argv[1:]))
