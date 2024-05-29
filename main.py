# main.py

# Importar módulos estándar
import time
import shutil
import win32com.client
from bs4 import BeautifulSoup
from tools import *
from io import StringIO
import re
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import NamedStyle, PatternFill, Border, Side, Font
from openpyxl.utils.dataframe import dataframe_to_rows
