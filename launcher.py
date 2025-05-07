# import streamlit
# import streamlit.web.cli as stcli
# import base64
# import os
# import sys
# from openpyxl import load_workbook, Workbook
# from openpyxl.styles import NamedStyle
# from copy import copy
# from datetime import datetime
# from io import BytesIO
# import pandas as pd
# from fpdf import FPDF
# from reportlab.pdfgen import canvas
# from reportlab.lib.pagesizes import A4
# from reportlab.lib.units import mm
# from reportlab.lib.colors import HexColor

# def resolve_path(path):
#     resolved_path = os.path.abspath(os.path.join(os.getcwd(), path))
#     return resolved_path


# if __name__ == "__main__":
#     sys.argv = [
#         "streamlit",
#         "run",
#         resolve_path("st.py"),
#         "--global.developmentMode=false",
#     ]
#     sys.exit(stcli.main())

import subprocess
import os

def resolve_path(path):
    return os.path.abspath(os.path.join(os.getcwd(), path))

if __name__ == "__main__":
    streamlit_script = resolve_path("st.py")  # Your main Streamlit file
    subprocess.Popen(
        ["streamlit", "run", streamlit_script, "--global.developmentMode=false"],
        creationflags=subprocess.CREATE_NO_WINDOW  # ✅ Hides console window
    )
