from setuptools import setup, find_packages

setup(
    name='excel2docx',
    version='0.1.0',
    author='Nabeel Kahlil Maulana',
    install_requires=[
        'et-xmlfile==1.1.0',
        'lxml==4.6.3',
        'openpyxl==3.0.7',
        'python-docx==0.8.11',
    ],
    author_email="nabeelkahli403@gmail.com",
    url='https://github.com/chawza/excel2docx',
    python_requires='>=3.7',
    packages=find_packages(),
    package_data = {
        'files' : ['*']
    },
    include_package_data=True
)