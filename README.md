# DIVA-PVDataManager

## Overview
This project is for managing ProjectDIVA's pv_db and auth_3d_db using Excel VBA. The code has been generated using Microsoft Copilot, with some minor adjustments.

## Requirements
- Microsoft Excel
- VBA enabled

*Note: This project has been tested with Excel 2013. Compatibility with other versions is unknown.*
*Note: Google Sheets does not support Excel VBA functionality, so please be aware of this limitation.*

## Workbook Descriptions
This project includes the following three types of Excel workbooks:

1. **Information Extraction Workbook** (DIVA_InfoExtract.xlsm):
   - Extracts and lists necessary information from `auth_3d_db.bin` and `pv_db`.

2. **a3da Data Management Workbook** (DIVA_a3daMgmt.xlsm):
   - Manages a3da data required for database registration in a list and writes it to `auth_3d_db.bin` as needed.

3. **pv_db Information Management Workbook** (DIVA_pvdbMgmt.xlsm):
   - Manages information included in `pv_db` (such as another song and `auth_3d` replacement information) and outputs it in `pv_db` format as needed.

## Usage
For detailed instructions on how to use DIVA-PVDataManager, please refer to the [Wiki](https://github.com/Riel2982/-DIVA-PVDataManager/wiki).

## Contributing
Please report bugs or request features via Issues or on Discord. Pull requests are also welcome. Feedback on compatibility with other versions is appreciated. Additionally, the development of standalone software that can achieve the same or better functionality is welcome.

### Note
Due to a lack of programming knowledge, I will make efforts to address issues, but I may not be able to respond to user requests. Contributions to improve and enhance the code are welcome.
