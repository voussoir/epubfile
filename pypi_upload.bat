py setup.py sdist
twine upload -r pypi dist\*
rmdir /s /q dist epubfile.egg-info
