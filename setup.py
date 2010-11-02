#!/usr/bin/env python

try:
    from setuptools import setup
except:
    from distutils.core import setup

setup(name='pyrtd',
      packages=['rtd'],
      version='0.1.0',
      description='Python client for Excel RealTimeData components.',
      long_description= \
          'Excel RealTimeData (RTD) components were introduced in Excel 2002. They\n'
          'provide a robust mechanism for inserting real-time data into a\n'
          'spreadsheet. Behined the scenes an RTD server is a COM object that\n'
          'implements the IRTDServer interface. pyrtd implements a Python RTD client\n'
          'that can receive real-time data from any RTD component.\n',
      author='James Brotchie',
      author_email='brotchie@gmail.com',
      url='http://code.google.com/p/pyrtd/',
      platforms=['win32'],
      license='Apache 2.0',
      classifiers=[
        'Development Status :: 4 - Beta',
        'License :: OSI Approved :: Apache Software License',
        'Intended Audience :: Developers',
        'Programming Language :: Python',
        'Operating System :: Microsoft :: Windows',
        'Topic :: Office/Business :: Financial',
        'Environment :: Win32 (MS Windows)',
        'Topic :: Software Development :: Libraries',
      ],
)
