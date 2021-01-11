# TableauDesktopPy
 Tools for extracting metadata from Tableau Desktop workbook files. This package parses the underlying xml of a workbook to retrieve information on its style and data connections.

 ## Installation

 Install with pip:

 ```pip install TableauDesktopPy```

 ## Usage

 1. Provide a valid Tableau workbook file ('.twb' or '.twbx') to declare a `Workbook` object:

 ```
import TableauDesktopPy as tdp

my_workbook = tdp.Workbook("C:\Users\bpewyllie\test_workbook.twbx")
 ```

 2. Then use any of the following methods to retrieve workbook metadata:

 * `Workbook.get_colors`
 * `Workbook.get_custom_sql`
 * `Workbook.get_excel`
 * `Workbook.get_fonts`
 * `Workbook.get_hidden_fields`
 * `Workbook.get_onedrive`
 * `Workbook.get_xml`

 3. Other metadata may be retrieved by calling the `Workbook.xml` attribute and parsing with an xml parser such as BeautifulSoup.


