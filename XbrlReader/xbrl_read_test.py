# xbrl から読み込む
from xbrl_proc import read_xbrl

xbrl_file = r"C:\local\users\nishimura\Programs\!Python\sandbox\Edinet\XBRL\jpcrp030000-asr-001_E00989-000_2018-03-31_01_2018-06-20.xbrl"
df = read_xbrl(xbrl_file)
