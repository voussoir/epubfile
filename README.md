epubfile
========

```Python
import epubfile
book = epubfile.Epub('mybook.epub')

for text_id in book.get_texts():
    soup = book.read_file(text_id, soup=True)
    ...
    book.write_file(text_id, soup)

for image_id in book.get_images():
    data = book.read_file(image_id)
    ...
    book.write_file(image_id, data)

# Note that this only reverses the spine in the opf file, it does not reverse
# the table of contents, since the toc is stored as nav.xhtml and toc.ncx.
book.set_spine_order(reversed(book.get_spine_order()))

cover_id = book.get_cover_image()
if cover_id:
    book.rename_file(cover_id, 'myfavoritecoverimage')

book.save('modifiedbook.epub')
```

epubfile provides simple editing of epub books. epubfile attempts to keep file modifications to a minimum. It does not add, remove, or rearrange files unless you ask it to, and does not inject additional metadata. As such, it works for both epub2 and epub3 assuming you stick to supported operations for your book version.

## Install

`pip install epubfile`

## Command line utilities

This library was born out of my own needs. So there are a couple of builtin utilities. You can run `epubfile.py` with no arguments to see a summary of all builtin utilities. Here are some examples:

```
addfile:
    Add files into the book.

covercomesfirst:
    Rename the cover image file so that it is the alphabetically-first image.

merge:
    Merge multiple books into one.

normalize:
    Rename files and directories in the book to match a common structure.
```

To see the rest, try `epubfile --help`.

## Spec compliance

epubfile does not rigorously enforce the epub spec and you can create noncompliant books with it. Basic errors are checked, and I am open to issues and comments regarding ways to improve spec-compliance without adding significant size or complexity to the library. I am prioritizing simplicity and ease of use over perfection.

## Pairs well with...

For advanced inter-file operations and better validation, I suggest using this library in conjunction with a good editor like [Sigil](https://github.com/Sigil-Ebook/Sigil). I wrote this library because although Sigil plugins are great for processing a single book, it is difficult to use Sigil to process multiple books, read book data for use in other programs, or do other inter-book operations.

## What not to expect

I do not intend to implement an object model for book metadata, beyond perhaps some basic getters and setters. You have full control over the `Epub.opf` BeautifulSoup object so you can edit the metadata however you want.

---

## Mirrors

https://git.voussoir.net/voussoir/epubfile

https://github.com/voussoir/epubfile

https://gitlab.com/voussoir/epubfile

https://codeberg.org/voussoir/epubfile
