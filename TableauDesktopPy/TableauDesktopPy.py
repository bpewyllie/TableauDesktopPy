from zipfile import ZipFile
from lxml import etree
import pandas as pd
import re


class Workbook:
    """
    Defines a workbook object from a filename.
    """

    def __init__(self, filename):
        self.filename = filename
        self.xml = self.get_xml()

    def get_xml(self):
        """
        Returns the xml of the given .twb or .twbx file.
        """

        # Ensure workbook_file is a workbook file
        twb = self.filename.split(".")

        if twb[-1][:3] != "twb" or len(twb) == 1:
            return self.filename + " is not a valid .twb or .twbx file."

        else:

            # unzip packaged workbooks to obtain xml
            if twb[-1] == "twb":
                xml = etree.parse(self.filename).getroot()

            else:
                with open(self.filename, "rb") as binfile:
                    twbx = ZipFile(binfile)
                    name = [w for w in twbx.namelist() if w.find(".twb") != -1][0]
                    unzip = twbx.open(name)
                    xml = etree.parse(unzip).getroot()

            return xml

    def get_custom_sql(self):
        """
        Returns a list of all unique custom sql queries in the workbook.
        """

        search = self.xml.xpath("//relation[@type='text']")
        queries = list(set([sql.text.lower() for sql in search]))

        return queries

    def get_excel(self):
        """
        Returns a list of excel and csv connections in the workbook.
        """

        search = self.xml.xpath("//connection[@filename != '']")
        files = list(set([xls.attrib["filename"].lower() for xls in search]))

        return files

    def get_onedrive(self):
        """
        Returns a list of onedrive connections in the workbook.
        """

        search = self.xml.xpath("//connection[@cloudFileProvider='onedrive']")
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

    def get_fonts(self):
        """
        Returns a list of fonts used in the workbook.
        """

        font_search = self.xml.xpath("//format[@attr = 'font-family']")
        fonts = list(set([font.attrib["value"].lower() for font in font_search]))

        return fonts

    def get_colors(self):
        """
        Returns dataframe of all individual colors and their associated elements in the workbook.
        """

        all_colors = []

        # Get worksheets in workbook
        wksht_search = self.xml.xpath("//worksheet")
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
            text_search = sheet.xpath(
                ".//formatted-text[./run[contains(@fontcolor, '#')]]"
            )
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

    def get_color_palettes(self):
        """
        Returns list of all named color palettes used in the workbook.
        """

        search = self.xml.xpath("//encoding[@palette !='']")
        palettes = list(set([color.attrib["palette"] for color in search]))

        return palettes

    def get_hidden_fields(self):
        """
        Returns list of all hidden fields in the workbook.
        """

        # return captions for calculated fields, otherwise return name
        has_caption = self.xml.xpath("//column[@caption and @hidden='true']")
        search = self.xml.xpath("//column[@hidden='true']")

        # replace brackets from field strings
        regex = r"^\[|\]\Z"

        hidden_fields = list(set([col.attrib["caption"] for col in has_caption]))
        hidden_fields += list(
            set(
                [
                    re.sub(regex, "", col.attrib["name"])
                    for col in search
                    if col not in has_caption
                ]
            )
        )

        return sorted(hidden_fields)

    def get_active_fields(self):
        """
        Returns list of all used fields in the workbook.
        """

        # return captions for calculated fields, otherwise return name
        has_caption = self.xml.xpath("//datasource-dependencies//column[@caption]")
        search = self.xml.xpath("//datasource-dependencies//column")

        # replace brackets from field strings
        regex = r"^\[|\]\Z"

        active_fields = list(set([col.attrib["caption"] for col in has_caption]))
        active_fields += list(
            set(
                [
                    re.sub(regex, "", col.attrib["name"])
                    for col in search
                    if col not in has_caption
                ]
            )
        )

        return sorted(active_fields)

    def get_fields(self):
        """
        Returns list of all fields in the workbook.
        """

        # return captions for calculated fields, otherwise return name
        has_caption = self.xml.xpath("//column[@caption and @name]")
        search = self.xml.xpath("//column[@name]")

        # replace brackets from field strings
        regex = r"^\[|\]\Z"

        fields = list(set([col.attrib["caption"] for col in has_caption]))
        fields += list(
            set(
                [
                    re.sub(regex, "", col.attrib["name"])
                    for col in search
                    if col not in has_caption
                ]
            )
        )

        return sorted(fields)

    def get_images(self):
        """
        Returns list of all image paths in the workbook.
        """

        search = self.xml.xpath(
            "//zone[@_.fcp.SetMembershipControl.false...type = 'bitmap']"
        )
        images = list(set([img.attrib["param"].lower() for img in search]))

        return images

    def get_shapes(self):
        """
        Returns list of all shape names in the workbook.
        """

        search = self.xml.xpath("//shape[@name != '']")
        shapes = list(set([shape.attrib["name"].lower() for shape in search]))

        return shapes