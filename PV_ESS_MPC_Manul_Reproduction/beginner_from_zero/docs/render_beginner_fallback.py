"""QA-only HTML rendering fallback for the beginner DOCX manual."""
from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path

BASE = Path(r"D:\PV_MPPT\PV_ESS_MPC_Paper_Reproduction\docs\render_docx_fallback.py")
spec = spec_from_file_location("docx_fallback_base", BASE)
module = module_from_spec(spec)
spec.loader.exec_module(module)

module.DOCX = Path(r"D:\PV_MPPT\PV_ESS_MPC_Paper_Reproduction\beginner_from_zero\docs\PV_ESS_MPC_From_Zero_Beginner_Manual.docx")
module.OUT = Path(r"D:\PV_MPPT\PV_ESS_MPC_Paper_Reproduction\beginner_from_zero\docs\qa_fallback_v2")
module.ASSETS = module.OUT / "assets"

if __name__ == "__main__":
    module.main()
