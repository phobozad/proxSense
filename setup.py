from distutils.core import setup
import py2exe

setup(
	windows=[
		{
			"script": 'proxSense.py',
			"icon_resources": [(1, "icon.ico")]
		}
	],
)