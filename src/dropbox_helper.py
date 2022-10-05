import yaml
import dropbox
import sys
from pprint import pprint as pp

with open("conf/dropbox.yaml", "r") as file:
    conf = yaml.safe_load(file)

dbx = dropbox.Dropbox(conf["token"])

res = dbx.files_list_folder(path="")

pp(res.entries)
sys.exit(0)
