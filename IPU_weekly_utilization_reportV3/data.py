# Duplicates :
#   - MR-IDMIICS-MRM-DEV
#   - MRAS-IDMIICS-HSB-DEV

# MasterOrgs :
#   - HistoricalTableData_EMEA
#   - HistoricalTableData_NA
#   - HistoricalTableData_EMEAPROD
#   - HistoricalTableData_MRSI
#   - HistoricalTableData_NAPROD

# Master Orgs
masterOrgs = ['HistoricalTableData_3zwXUNlxK28', 'HistoricalTableData_9e2PhXBxLfC', 'HistoricalTableData_l19jxOq26eZ', 'HistoricalTableData_cGocp7Btyqm', 'HistoricalTableData_86rkOFSaWMr']

# Special Sheets - Summary Sheet Name, Souce Master Org, Source Sub Orgs(list)
specialOrgs = [
    ('NA-AMIG - Summary Reports', 'HistoricalTableData_NAAMIG', ['HistoricalTableData_9e2PhXBxLfC', 'HistoricalTableData_86rkOFSaWMr'], ['MRAS-IDMIICS-AMIG-DEV', 'MRAS-IDMIICS-AMIG-QA', 'MRAS-IDMIICS-AMIG-PROD', 'Sub-Organization']),
    ('NA-MRSI - Summary Reports', 'HistoricalTableData_NAMRSI', ['HistoricalTableData_9e2PhXBxLfC'], ['MRAS_IDMIICS_MRSI_DEV', 'MRAS_IDMIICS_MRSI_QA', 'MRAS_IDMIICS_MRSI_UAT', 'Sub-Organization']),
    ('MRSI - Summary (main)', 'HistoricalTableData_MRSI', ['HistoricalTableData_cGocp7Btyqm'], ['MRAS-IDMIICS-DEV', 'MRAS-IDMIICS-PRD', 'MRAS-IDMIICS-QA', 'MRAS-IDMIICS-SIT', 'MRAS-IDMIICS-UAT', 'Sub-Organization']),
    ('HSB - Summary Reports', 'HistoricalTableData_HSB' , ['HistoricalTableData_cGocp7Btyqm'], ['MRAS-IDMIICS-HSB-DEV', 'MRAS-IDMIICS-HSB-UAT'])
]

# Wihtout Region
orgList = {'HistoricalTableData_3zwXUNlxK28': 'HistoricalTableData_EMEA', 'HistoricalTableData_lstdiXKyQE5': 'MR-IDMIICS-GLISE-DEV', 'HistoricalTableData_2B83FrdeI6m': 'MR-IDMIICS-GLISE-SIT', 'HistoricalTableData_5L2IEkXMmuq': 'MR-IDMIICS-GLISE-UAT', 'HistoricalTableData_8lkzBtWxEca': 'MR-IDMIICS-MRM-DEV', 'HistoricalTableData_bSxbioBLXiz': 'MR-IDMIICS-MRM-UAT', 'HistoricalTableData_c9vyhnDmTy6': 'MR-IDMIICS-GIM-DEV', 'HistoricalTableData_9e2PhXBxLfC': 'HistoricalTableData_NA', 'HistoricalTableData_fK1q3wgm5oP': 'NA_MR-IDMIICS-MRM-DEV', 'HistoricalTableData_90mWdtGoP7K': 'MRAS-IDMIICS-AMIG-DEV', 'HistoricalTableData_dWc4Zw5TfOq': 'MRAS-IDMIICS-AMIG-QA', 'HistoricalTableData_cmURWkN3gO6': 'MRAS-IDMIICS-COE-SBX', 'HistoricalTableData_50ivoX5vG8m': 'NA_MRAS-IDMIICS-HSB-DEV', 'HistoricalTableData_2dg7AUJCzyh': 'MRAS_IDMIICS_MRSI_DEV', 'HistoricalTableData_7ZGshrDuK2G': 'MRAS_IDMIICS_MRSI_QA', 'HistoricalTableData_8728Mf2aqow': 'MRAS_IDMIICS_MRSI_UAT', 'HistoricalTableData_l19jxOq26eZ': 'HistoricalTableData_EMEAPROD', 'HistoricalTableData_dRkDPPHfC40': 'MR-IDMIICS-GLISE-PRD', 'HistoricalTableData_75zoCBUNCiG': 'MR-IDMIICS-GLISE-PREPRD', 'HistoricalTableData_3oNWmIBysqX': 'MR-IDMIICS-MRM-PRD', 'HistoricalTableData_cGocp7Btyqm': 'HistoricalTableData_MRSI', 'HistoricalTableData_bLwgJRNBpwT': 'MRAS-IDM-AMIG-DEV', 'HistoricalTableData_1BDAyhrwb2E': 'MRAS-IDMIICS-DEV', 'HistoricalTableData_2RFFcLuMGys': 'MRAS-IDMIICS-DEV2', 'HistoricalTableData_1RzdpkquZaY': 'MRAS-IDMIICS-DEVINT', 'HistoricalTableData_kvQJDn0xV0g': 'MRAS-IDMIICS-HSB-DEV', 'HistoricalTableData_0P60suZWD3X': 'MRAS-IDMIICS-HSB-QA', 'HistoricalTableData_000SF39jeVm': 'MRAS-IDMIICS-HSB-SIT', 'HistoricalTableData_6g2PSRCeDYs': 'MRAS-IDMIICS-HSB-UAT', 'HistoricalTableData_3v0NIUXqhxL': 'MRAS-IDMIICS-MRSI-DEV', 'HistoricalTableData_2pLJ8Pi6bq8': 'MRAS-IDMIICS-POC', 'HistoricalTableData_1GMTUrbdefb': 'MRAS-IDMIICS-PRD', 'HistoricalTableData_9bsf7eJZiEt': 'MRAS-IDMIICS-QA', 'HistoricalTableData_gwHM5AIfVkF': 'MRAS-IDMIICS-SIT', 'HistoricalTableData_8xXabpjlW9Y': 'MRAS-IDMIICS-UAT', 'HistoricalTableData_1VYmnjJ71Ut': 'MRAS-IDMIICS-UAT2', 'HistoricalTableData_86rkOFSaWMr': 'HistoricalTableData_NAPROD', 'HistoricalTableData_es31IPGwBSD': 'MRAS-IDMIICS-AMIG-PROD', 'HistoricalTableData_gbI3gTKHMmm':'MRAS_IDMIICS_MRSI_PROD'}

# Filenames
filenamesToSave = {'HistoricalTableData_3zwXUNlxK28': 'HistoricalTableData_EMEA', 'HistoricalTableData_lstdiXKyQE5': 'EMEA-GLISE-DEV', 'HistoricalTableData_2B83FrdeI6m': 'EMEA-GLISE-SIT', 'HistoricalTableData_5L2IEkXMmuq': 'EMEA-GLISE-UAT', 'HistoricalTableData_8lkzBtWxEca': 'EMEA-MRM-DEV', 'HistoricalTableData_bSxbioBLXiz': 'EMEA-MRM-UAT', 'HistoricalTableData_c9vyhnDmTy6': 'EMEA-GIM-DEV', 'HistoricalTableData_9e2PhXBxLfC': 'HistoricalTableData_NA', 'HistoricalTableData_fK1q3wgm5oP': 'NA_MR-IDMIICS-MRM-DEV', 'HistoricalTableData_90mWdtGoP7K': 'NA-AMIG-DEV', 'HistoricalTableData_dWc4Zw5TfOq': 'NA-AMIG-QA', 'HistoricalTableData_cmURWkN3gO6': 'NA-COE-SBX', 'HistoricalTableData_50ivoX5vG8m': 'NA_MRAS-IDMIICS-HSB-DEV', 'HistoricalTableData_2dg7AUJCzyh': 'NA-MRSI-DEV', 'HistoricalTableData_7ZGshrDuK2G': 'NA-MRSI-QA', 'HistoricalTableData_8728Mf2aqow': 'NA-MRSI-UAT', 'HistoricalTableData_l19jxOq26eZ': 'HistoricalTableData_EMEAPROD', 'HistoricalTableData_dRkDPPHfC40': 'EMEAPROD-GLISE-PRD', 'HistoricalTableData_75zoCBUNCiG': 'EMEAPROD-GLISE-PREPRD', 'HistoricalTableData_3oNWmIBysqX': 'EMEAPROD-MRM-PRD', 'HistoricalTableData_cGocp7Btyqm': 'HistoricalTableData_MRSI', 'HistoricalTableData_bLwgJRNBpwT': 'MRAS-IDM-AMIG-DEV', 'HistoricalTableData_1BDAyhrwb2E': 'MRSI-DEV', 'HistoricalTableData_2RFFcLuMGys': 'MRSI-DEV2', 'HistoricalTableData_1RzdpkquZaY': 'MRSI-DEVINT', 'HistoricalTableData_kvQJDn0xV0g': 'MRSI-HSB-DEV', 'HistoricalTableData_0P60suZWD3X': 'MRSI-HSB-QA', 'HistoricalTableData_000SF39jeVm': 'MRSI-HSB-SIT', 'HistoricalTableData_6g2PSRCeDYs': 'MRSI-HSB-UAT', 'HistoricalTableData_3v0NIUXqhxL': 'MRSI-DEV', 'HistoricalTableData_2pLJ8Pi6bq8': 'MRSI-POC', 'HistoricalTableData_1GMTUrbdefb': 'MRSI-PRD', 'HistoricalTableData_9bsf7eJZiEt': 'MRSI-QA', 'HistoricalTableData_gwHM5AIfVkF': 'MRSI-SIT', 'HistoricalTableData_8xXabpjlW9Y': 'MRSI-UAT', 'HistoricalTableData_1VYmnjJ71Ut': 'MRSI-UAT2', 'HistoricalTableData_86rkOFSaWMr': 'HistoricalTableData_NAPROD', 'HistoricalTableData_es31IPGwBSD': 'NAPROD-AMIG-PROD', 'HistoricalTableData_gbI3gTKHMmm':'NAPROD-MRSI-PROD'}

# Orgs without IPU data
exceptionOrgs = ['MRAS-IDM-AMIG-DEV', 'Sub-Organization']

# MMDDYYYY
date = "04192024"

# Column names for Excel data
header = ['Date', 'Month', 'Meter', 'Utilization', 'IPUs Consumed']

sourcePath = "C:/Users/lpenke/OneDrive - Informatica/Desktop/MunichRE/Automations/IPU_weekly_utilization_report/source/"
sourcePath = "C:/Users/lpenke/OneDrive - Informatica/Desktop/MunichRE/Automations/FileRename/ParsedFiles/"
targetPath = "C:/Users/lpenke/OneDrive - Informatica/Desktop/MunichRE/Automations/IPU_weekly_utilization_report/target/"

subOrgConsumptionFileName = 'sub-organization_consumption.xlsx'

tgtFilename = 'MRE IDMC IPU Report '+ date +'.xlsx'
