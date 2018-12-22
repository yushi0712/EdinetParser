# xbrl から読み込む
from xbrl_proc import read_xbrl

xbrl_file = r"C:/local/users/nishimura/Documents/!Theme/財務分析/XBRL/E00989_コニカミノルタ株式会社/jpcrp030000-asr-001_E00989-000_2018-03-31_01_2018-06-20.xbrl"
df = read_xbrl(xbrl_file)
