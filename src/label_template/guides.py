import os
import zipfile
import xml.etree.ElementTree as ET
from pptx.oxml.ns import qn, _nsmap

import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils import qn_xpath, pretty_print_xml

# p15 네임스페이스 추가
_nsmap['p15'] = 'http://schemas.microsoft.com/office/powerpoint/2012/main'


class PowerPointGuideEditor:
    def __init__(self, pptx_file: str):
        self.pptx_file = pptx_file
        self.temp_dir = "pptx_temp"
        self.view_props_path = os.path.join(self.temp_dir, "ppt/viewProps.xml")
        self.presentation_path = os.path.join(self.temp_dir, "ppt/presentation.xml")
        
        self._extract_pptx()
        self._load_xml_trees()
    
    def _extract_pptx(self):
        """PPTX 파일을 압축 해제합니다."""
        with zipfile.ZipFile(self.pptx_file, "r") as zip_ref:
            zip_ref.extractall(self.temp_dir)
    
    def _load_xml_trees(self):
        """XML 트리를 로드합니다."""
        self.view_props_tree = ET.parse(self.view_props_path) if os.path.exists(self.view_props_path) else None
        self.presentation_tree = ET.parse(self.presentation_path) if os.path.exists(self.presentation_path) else None
        assert self.view_props_tree or self.presentation_tree, "PPTX 파일을 불러오는 데 실패했습니다."
        self.view_props_root = self.view_props_tree.getroot() if self.view_props_tree else None
        self.presentation_root = self.presentation_tree.getroot() if self.presentation_tree else None
        for ns in _nsmap:
            ET.register_namespace(ns, _nsmap[ns])
                
    def _save_pptx(self, output_pptx: str):
        """변경된 XML을 압축하여 새로운 PPTX 파일로 저장합니다."""
        with zipfile.ZipFile(output_pptx, "w", zipfile.ZIP_DEFLATED) as new_zip:
            for folder, _, files in os.walk(self.temp_dir):
                for file in files:
                    file_path = os.path.join(folder, file)
                    arcname = file_path.replace(self.temp_dir + os.sep, "")  # ZIP 내부 경로 유지
                    new_zip.write(file_path, arcname)
    
    def _cleanup_temp(self):
        """임시 폴더 삭제"""
        for root, _, files in os.walk(self.temp_dir, topdown=False):
            for file in files:
                os.remove(os.path.join(root, file))
            os.rmdir(root)

    def add_guide(self, pos: int, orient: str = None, color: str = None):
        """
        개별 안내선을 추가하는 함수. ID를 자동으로 할당하며, 위치와 색상을 개별적으로 추가 가능.
        
        :param pos: 안내선 위치
        :param orient: 'horz' (수평) 또는 None (수직)
        :param color: 안내선 색상 (RGB HEX 코드)
        """
        # 🔹 viewProps.xml 수정 (위치 추가)
        guide_list = self.view_props_root.find(qn_xpath(".//p:slideViewPr/p:cSldViewPr/p:guideLst"))
        if guide_list is None:
            guide_list = ET.SubElement(self.view_props_root.find(qn_xpath(".//p:slideViewPr/p:cSldViewPr")), qn("p:guideLst"))
        
        ET.SubElement(guide_list, qn("p:guide"), {"pos": str(pos), **({"orient": orient} if orient else {})})

        # 🔹 presentation.xml 수정 (색상 추가)
        sldGuideLst = self.presentation_root.find(qn_xpath(".//p:extLst/p:ext/p15:sldGuideLst"))
        if sldGuideLst is None:
            sldGuideLst = ET.SubElement(self.presentation_root.find(qn_xpath(".//p:extLst/p:ext")), qn("p15:sldGuideLst"))
        
        guide_id = len(sldGuideLst.findall(qn("p15:guide"))) + 1  # 자동 ID 할당
        guide = ET.SubElement(sldGuideLst, qn("p15:guide"), {"id": str(guide_id), "pos": str(pos), **({"orient": orient} if orient else {}), "userDrawn": "1"})
        
        if color:
            clr_elem = ET.SubElement(guide, qn("p15:clr"))
            ET.SubElement(clr_elem, qn("a:srgbClr"), {"val": color})
    
    def write(self):
        self.view_props_tree.write(self.view_props_path, xml_declaration=True, encoding="UTF-8")
        self.presentation_tree.write(self.presentation_path, xml_declaration=True, encoding="UTF-8")

    def save(self, output_pptx: str):
        self.write()
        self._save_pptx(output_pptx)
        self._cleanup_temp()

if __name__ == "__main__":
    # $env:PYTHONPATH = "$PWD\src" && python '.\src\label_template\guides.py' # 잘 안됨...
    input_pptx = "template/nametag.pptx"
    output_pptx = "example/ppt_with_guides.pptx"

    editor = PowerPointGuideEditor(input_pptx)
    editor.add_guide(2160, None, "FF0000")
    editor.add_guide(1487, "horz", "00FF00")
    editor.add_guide(4458, "horz", "0000FF")
    editor.add_guide(3181, None, "FFA500")
    editor.save(output_pptx)
