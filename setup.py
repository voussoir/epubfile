import setuptools

setuptools.setup(
    name='epubfile',
    py_modules=['epubfile'],
    version='0.0.1',
    author='voussoir',
    author_email='ethan@voussoir.net',
    description='simple epub file reading and writing',
    long_description=open('README.md', 'r').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/voussoir/epubfile',
    install_requires=['bs4', 'html5lib', 'tinycss2', 'voussoirkit'],
)
