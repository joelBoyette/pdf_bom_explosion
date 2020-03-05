
from excel_access_run_macro import run_excel_macro, run_access_macro


def refresh_data():

    # update part data for traverse bom file
    run_access_macro(access_path=r'\\vfile\MPPublic\Kanban Projects\Kanban.accdb',
                     macros=['update_oh,ono,dmd_macro'])

    run_excel_macro(excel_path=r'\\vfile\MPPublic\ECN Status\ecn_data.xlsm',
                    macros=['refresh_epicor_data'])

    # update bin data for bin locations file
    run_excel_macro(r'\\vfile\MPPublic\Projects\NPD\NPD_data.xlsm', macros=['refresh_bin_contents'])