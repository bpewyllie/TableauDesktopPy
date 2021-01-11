from zipfile import ZipFile
from lxml import etree
import pandas as pd


def get_xml(workbook_file):

    """
    Returns the xml of the given .twb or .twbx file.
    """

    # Ensure workbook_file is a workbook file
    twb = workbook_file.split(".")

    if twb[-1][:3] != "twb" or len(twb) == 1:
        return workbook_file + " is not a valid .twb or .twbx file."

    else:

        # unzip packaged workbooks to obtain xml
        if twb[-1] == "twb":
            xml = etree.parse(workbook_file).getroot()

        else:
            with open(workbook_file, "rb") as binfile:
                twbx = ZipFile(binfile)
                name = [w for w in twbx.namelist() if w.find(".twb") != -1][0]
                unzip = twbx.open(name)
                xml = etree.parse(unzip).getroot()

        return xml


def check_custom_sql(xml):

    """
    Returns number of custom sql queries in the workbook.
    """

    search = xml.xpath("//relation[@type='text']")
    n_queries = len(list(set([sql.text.lower() for sql in search])))

    return n_queries


def check_excel(xml):

    """
    Returns a list of excel and csv connections in the workbook.
    """

    search = xml.xpath("//connection[@filename != '']")
    files = list(set([xls.attrib["filename"].lower() for xls in search]))

    return files


def check_onedrive(xml):

    """
    Returns a list of onedrive connections in the workbook.
    """

    search = xml.xpath("//connection[@cloudFileProvider='onedrive']")
    onedrive = list(set([od.attrib["filename"].lower() for od in search]))

    return onedrive

    # def check_db_connections(xml):

    #     """
    #     Returns a list of other database connections in the workbook.
    #     """

    #     search = xml.xpath("//connection[@dbname != '']")
        # dbs = list(
        #     set(
        #         [
        #             (
        #                 db.attrib["dbname"],
        #                 db.xpath(".//connection[@auto-extract !='']").attrib[
        #                     "auto-extract"
        #                 ],
        #             )
        #             for db in search
        #         ]
        #     )
        # )


#     return dbs


# def check_permissions(xml):

#     """
#     Checks that the user has modified the default permissions for the workbook.
#     """


def check_fonts(xml):

    """
    Returns fonts used in the workbook.
    """

    font_search = xml.xpath("//format[@attr = 'font-family']")
    fonts = list(set([font.attrib["value"].lower() for font in font_search]))

    return fonts


def check_colors(xml):

    """
    Returns colors used in the workbook.
    """

    all_colors = []

    # Get worksheets in workbook
    wksht_search = xml.xpath("//worksheet")
    for sheet in wksht_search:

        # Get style elements for each worksheet
        style_search = sheet.xpath(".//style-rule[./format[contains(@value, '#')]]")
        for style in style_search:

            # Get colors for each style element
            color_search = style.xpath(".//format[contains(@value, '#')]")
            for color in color_search:
                all_colors.append(
                    (
                        sheet.attrib["name"],
                        style.attrib["element"],
                        color.attrib["value"],
                    )
                )

        # Get tooltip elements for each worksheet
        text_search = sheet.xpath(".//formatted-text[./run[contains(@fontcolor, '#')]]")
        for text in text_search:

            # Get fontcolors for each tooltip text element
            ttcolor_search = text.xpath(".//run[contains(@fontcolor, '#')]")
            for color in ttcolor_search:
                all_colors.append(
                    (
                        sheet.attrib["name"],
                        "tooltip",
                        color.attrib["fontcolor"],
                    )
                )

    unique_colors = list(set(all_colors))

    color_df = pd.DataFrame(
        unique_colors, columns=["Sheet", "Element", "Color"]
    ).sort_values(["Sheet", "Element"])

    return color_df


def check_fields(xml):

    """
    Returns list of all hidden fields in the workbook.
    """

    search = xml.xpath("//column[@hidden='true']")
    hidden_fields = list(set([col.attrib["name"] for col in search]))

    return hidden_fields
