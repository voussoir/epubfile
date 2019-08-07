epubfile
========

```Python
import epubfile
book = epubfile.Epub.open('mybook.epub')

for text_id in book.get_texts():
    soup = book.read_file(text_id, soup=True)
    ...
    book.write_file(text_id, soup)

for image_id in book.get_images():
    data = book.read_file(image_id)
    ...
    book.write_file(image_id, data)

# Note that this does not reverse the table of contents.
book.set_spine_order(reversed(book.get_spine_order()))

cover_id = book.get_cover_image()
if cover_id:
    book.rename_file(cover_id, 'myfavoritecoverimage')

book.save('modifiedbook.epub')
```

epubfile provides simple editing of epub books. epubfile attempts to keep file modifications to a minimum. It does not add, remove, or rearrange files unless you ask it to, and does not inject additional metadata. As such, it works for both epub2 and epub3 assuming you stick to supported operations for your book version.

# Install

`pip install epubfile`

# Command line utilities

This library was born out of my own needs. So there are a couple of builtin utilities.

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

# Spec compliance

epubfile does not rigorously enforce the epub spec and you can create noncompliant books with it. Basic errors are checked, and I am open to issues and comments regarding ways to improve spec-compliance without adding significant size or complexity to the library. I am prioritizing simplicity and ease of use over perfection.

# Pairs well with...

For advanced inter-file operations and better validation, I suggest using this library in conjunction with a good editor like [Sigil](https://github.com/Sigil-Ebook/Sigil). I wrote this library because although Sigil plugins are great for processing a single book, it is difficult to use Sigil to process multiple books, read book data for use in other programs, or do other inter-book operations.

# What not to expect

I do not intend to implement an object model for book metadata, beyond perhaps some basic getters and setters. You have full control over the `Epub.opf` BeautifulSoup object so you can edit the metadata however you want.

---

```
BSD 3-Clause License

Copyright (c) 2019, Ethan Dalool
https://github.com/voussoir/epubfile
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its
   contributors may be used to endorse or promote products derived from
   this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
```
