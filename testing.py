#
#
# def __win32_finddll():
#     from winreg import (
#         OpenKey,
#         CloseKey,
#         EnumKey,
#         QueryValueEx,
#         QueryInfoKey,
#         HKEY_LOCAL_MACHINE,
#     )
#
#     from distutils.version import LooseVersion
#     import os
#
#     dlls = []
#     # Look up different variants of Ghostscript and take the highest
#     # version for which the DLL is to be found in the filesystem.
#     for key_name in (
#         "AFPL Ghostscript",
#         "Aladdin Ghostscript",
#         "GNU Ghostscript",
#         "GPL Ghostscript",
#     ):
#         print(f'local machine key {HKEY_LOCAL_MACHINE}')
#         try:
#             print("Software\\%s" % key_name)
#             k1 = OpenKey(HKEY_LOCAL_MACHINE, "Software\\%s" % key_name)
#
#             for num in range(0, QueryInfoKey(k1)[0]):
#                 print(QueryInfoKey(k1))
#                 version = EnumKey(k1, num)
#                 print(version)
#                 try:
#                     k2 = OpenKey(k1, version)
#                     dll_path = QueryValueEx(k2, "GS_DLL")[0]
#                     CloseKey(k2)
#                     if os.path.exists(dll_path):
#                         dlls.append((LooseVersion(version), dll_path))
#                 except WindowsError:
#                     pass
#             CloseKey(k1)
#         except WindowsError:
#             pass
#     if dlls:
#         dlls.sort()
#         return dlls[-1][-1]
#     else:
#         return None
#
# tttt = __win32_finddll()
# print(tttt)

import pandas as pd

epicor_data = pd.read_excel(r'C:\Users\JBoyette.BRLEE\Documents\new_scripts\test_epicor_data.xlsx',
                            sheet_name='TEST')

df = epicor_data.loc[epicor_data['Phantom BOM']]

print(df)