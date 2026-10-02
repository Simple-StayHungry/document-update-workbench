import io,tempfile,unittest,zipfile
from pathlib import Path
from docx import Document as PyDoc
import sys
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from server import prepare_materials

def docx_bytes(lines):
    d=PyDoc()
    for x in lines:d.add_paragraph(x)
    f=io.BytesIO();d.save(f);return f.getvalue()

class IntakeTests(unittest.TestCase):
    def test_one_box_classifies_source_and_target_zip(self):
        src=docx_bytes(['河南测试有限公司','2026年面向专业投资者公开发行公司债券','募集说明书','第一节 总则'])
        tar=docx_bytes(['控股股东情况说明','公司说明如下：','具体情况详见募集说明书。'])
        zb=io.BytesIO()
        with zipfile.ZipFile(zb,'w') as z:z.writestr('说明性文件/目标.docx',tar)
        with tempfile.TemporaryDirectory() as td:
            source,targets,meta=prepare_materials(Path(td),[('最新募集说明书.docx',src),('说明性文件.zip',zb.getvalue())])
            self.assertEqual(source['name'],'最新募集说明书.docx');self.assertEqual([x[1] for x in targets],['目标.docx']);self.assertEqual(meta['targets'],1)
    def test_target_body_reference_is_not_mistaken_for_prospectus(self):
        src=docx_bytes(['河南测试有限公司','2026年面向专业投资者公开发行公司债券','募集说明书'])
        tar=docx_bytes(['募集资金情况说明','公司本次使用安排与募集说明书约定一致。'])
        with tempfile.TemporaryDirectory() as td:
            source,targets,_=prepare_materials(Path(td),[('A.docx',src),('说明.docx',tar)])
            self.assertEqual(source['name'],'A.docx');self.assertEqual(targets[0][1],'说明.docx')
    def test_duplicate_docx_is_deduped(self):
        src=docx_bytes(['河南测试有限公司','2026年面向专业投资者公开发行公司债券','募集说明书'])
        tar=docx_bytes(['情况说明','公司说明如下。'])
        with tempfile.TemporaryDirectory() as td:
            _,targets,meta=prepare_materials(Path(td),[('募集说明书.docx',src),('a.docx',tar),('copy.docx',tar)])
            self.assertEqual(len(targets),1);self.assertEqual(meta['docx_found'],2)
    def test_multiple_prospectuses_are_blocked(self):
        a=docx_bytes(['甲公司','2026年公司债券','募集说明书']);b=docx_bytes(['乙公司','2026年公司债券','募集说明书']);tar=docx_bytes(['情况说明'])
        with tempfile.TemporaryDirectory() as td:
            with self.assertRaisesRegex(ValueError,'多个募集说明书'):prepare_materials(Path(td),[('募集说明书A.docx',a),('募集说明书B.docx',b),('说明.docx',tar)])
if __name__=='__main__':unittest.main()
