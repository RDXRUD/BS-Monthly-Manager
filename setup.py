from cx_Freeze import setup, Executable

setup(
    name="Monthly Manager",
    version="1.0",
    description="Streamlit without console",
    executables=[
        Executable("launcher.py", base="Win32GUI")  # Use your actual launcher filename
    ]
)
