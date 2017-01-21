from distutils.core import setup
setup(
    name = 'pypowerworld',
    author = 'Jesse Boyd',
    author_email = 'jessemarkboyd@gmail.com',
    url = 'https://github.com/jessemarkboyd/pypowerworld',
    download_url = 'https://github.com/jessemarkboyd/pypowerworld/tarball/0.2.3',
    keywords = ['testing', 'logging', 'powerworld', 'powerflow', 'loadflow'],
    version = '0.2.3',
    description = 'Powerworld COM wrapper for Python',
    long_description = open('README.txt',encoding='utf8').read(),
    packages = ['pypowerworld'],
    license = 'Creative Commons Attribution-Noncommercial-Share Alike license',
)