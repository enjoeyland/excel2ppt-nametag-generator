from pptx.oxml import register_element_cls
from pptx.oxml.text import CT_TextCharacterProperties, CT_TextFont
from pptx.oxml.xmlchemy import ZeroOrOne
from pptx.text.text import Font, _Run

# Support for east asian font which is not implemented in python-pptx.
ea = ZeroOrOne(
        "a:ea",
        successors=(
            "a:ea",
            "a:cs",
            "a:sym",
            "a:hlinkClick",
            "a:hlinkMouseOver",
            "a:rtl",
            "a:extLst",
        ),
    )

setattr(CT_TextCharacterProperties, 'ea', property(ea))
ea.populate_class_members(CT_TextCharacterProperties, "ea")
register_element_cls('a:ea', CT_TextFont)


def name(self, value):
    if value is None:
        self._rPr._remove_latin()
        self._rPr._remove_ea()
    else:
        latin = self._rPr.get_or_add_latin()
        latin.typeface = value
        ea = self._rPr.get_or_add_ea()
        ea.typeface = value

Font.name = Font.name.setter(name)