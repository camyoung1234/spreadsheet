# Copyright (c) 2010-2011 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file
import sys

from spreadsheet.openpyxl.shared.compat.elementtree import iterparse
from spreadsheet.openpyxl.shared.compat.tempnamedfile import NamedTemporaryFile
from spreadsheet.openpyxl.shared.compat.allany import all, any
from spreadsheet.openpyxl.shared.compat.strings import basestring, unicode, StringIO, file, BytesIO
from spreadsheet.openpyxl.shared.compat.numbers import long
from spreadsheet.openpyxl.shared.compat.itertools import ifilter, xrange

try:
    from collections import OrderedDict
except ImportError:
    from spreadsheet.openpyxl.shared.compat.odict import OrderedDict
