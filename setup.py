from setuptools import setup

setup(
    name='Outlook Signature Smuggler',
    version='0.1',
    py_modules=['sigsmuggle'],
    install_requires=[
        'Click',
    ],
    entry_points='''
        [console_scripts]
        sigsmuggle=sigsmuggle:cli
    ''',
)
