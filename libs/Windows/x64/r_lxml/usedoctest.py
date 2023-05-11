"""Doctest module for XML comparison.

Usage::

   >>> import r_lxml.usedoctest
   >>> # now do your XML doctests ...

See `lxml.doctestcompare`
"""

from r_lxml import doctestcompare

doctestcompare.temp_install(del_module=__name__)
