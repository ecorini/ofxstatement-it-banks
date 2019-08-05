~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ofxstatement-it-banks
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

`ofxstatement`_ is a tool to convert proprietary bank statement to OFX format,
suitable for importing to `GnuCash`_, `HomeBank`_ or any other standard-friendly accounting sofware.

`ofxstatement-it-banks`_ provides some italian banks plugins for ofxstatement.

Supported banks:

* Widiba (https://www.widiba.it) plugin 'widiba'
* WeBank (https://www.webank.it) plugin 'webank'

Most of the code was taken from projects `ofxstatement-fineco`_ and `ofxstatement-intesasp`_.



Installation
============

You can install the plugin as usual from pip or directly from the downloaded git

pip
---

::

  pip3 install --user ofxstatement-it-banks

setup.py
--------

::

  python3 setup.py install --user

Configuration
===============================
To edit the config file run this command
::

  $ ofxstatement edit-config


for example:
::

        [widiba]
        bank_id = Widiba
        causale2name = True
        plugin = widiba

        [webank]
        bank_id = Webank
        account_id = 00000 - 0000000000
        plugin = webank



Save and exit the text editor

configuration parameters
------------------------

widiba
--------

``bank``
        Bank id
        (default is 'Widiba')
``currency``
        Currency
        (default is 'EUR')
``account_type``
        Account type
        (default is 'CHECKING')
``info2name``
        If set to ``True`` copy content of column ``CAUSALE`` in xlsx file to ``<NAME>`` field in ofx file.
        Useful if ofx file is imported in `HomeBank`_, to populate 'info' field.
``info2memo``
        If set to ``True`` concat content of column ``CAUSALE`` in xlsx file at the end of the ``<MEMO>`` field in ofx file.
        Useful if ofx file is imported in `HomeBank`_, to setup automatic assignments rules.

webank
--------

``bank``
        Bank id
        (default is 'Webank')
``currency``
        Currency
        (default is 'EUR')
``account_id``
        Account id
        (default is '00000 - 0000000000')        
``account_type``
        Account type
        (default is 'CHECKING')

about the parsing
-----------------
- A semi-unique id is generated for any transaction through a ofxstatement's native method.
- The ofx transaction type are set to generic DEBIT or CREDIT according to the sign of transaction.
- The full description available is set to the memo field.


The plugin only support the savings statements xlsx file (no credit card support).

.. _ofxstatement: https://github.com/kedder/ofxstatement
.. _ofxstatement-fineco: https://github.com/frankIT/ofxstatement-fineco
.. _ofxstatement-intesasp: https://github.com/Jacotsu/ofxstatement-intesasp
.. _GnuCash: https://www.gnucash.org/
.. _HomeBank: http://homebank.free.fr/
.. _HomeBank limit: https://bugs.launchpad.net/homebank/+bug/1645124
