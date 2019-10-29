# coding=utf-8
import logging
import sys
import unittest

from exchangelib.util import PrettyXmlHandler

# Always show full repr() output for object instances in unittest error messages
unittest.util._MAX_LENGTH = 2000

argv = sys.argv.copy()
if '-v' in argv:
    logging.basicConfig(level=logging.DEBUG, handlers=[PrettyXmlHandler()])
else:
    logging.basicConfig(level=logging.CRITICAL)
