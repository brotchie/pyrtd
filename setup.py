#!/usr/bin/env python

try:
    from setuptools import setup
except:
    from distutils.core import setup

setup(name='pyrtd',
      packages=['rtd'],
      version='0.1.0',
      description='Python client for Excel RealTimeData components.',
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
