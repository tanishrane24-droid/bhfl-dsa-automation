from setuptools import setup
import sys
from pathlib import Path

# Read requirements
with open('requirements.txt') as f:
    requirements = [line.strip() for line in f if line.strip() and not line.startswith('#')]

setup(
    name='BHFL_DSA_Automation',
    version='1.0.0',
    description='BHFL DSA Automation - One-click desktop application for DSA mail distribution',
    author='BHFL',
    python_requires='>=3.8',
    install_requires=requirements,
    entry_points={
        'console_scripts': [
            'bhfl-dsa=main:main',
        ],
    },
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Intended Audience :: End Users/Desktop',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Operating System :: Microsoft :: Windows',
    ],
)
