import os
from gus_gas_common import update_master_excel

YEAR = int(os.getenv("GUS_YEAR", "2019"))

update_master_excel(year=YEAR, create_new=False)
