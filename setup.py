from distutils.core import setup

setup(
	name='recordlib',
	version='0.0.8',
	description='Simple records type of dict-list parser',
	author = 'Moon, Heung-sub',
	author_email = 'mhs9089@gmail.com',
	py_modules = ['recordlib'],
	install_requires=['xlrd', 'xlsxwriter'],
)