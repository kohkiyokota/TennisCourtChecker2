import os
from os.path import join, dirname

from dotenv import load_dotenv
dotenv_path = join(dirname(__file__), '.env')
load_dotenv(dotenv_path)

# Herokuで実行時では.envがないので上記コードでは何も起こらない
# .envから環境変数を読み込み
LINE_NOTIFY_TOKEN = os.environ['LINE_NOTIFY_TOKEN']
LNT_FOR_ERROR = os.environ['LNT_FOR_ERROR']

SUN=os.environ['SUN'].replace(' ', '')
MON=os.environ['MON'].replace(' ', '')
TUE=os.environ['TUE'].replace(' ', '')
WED=os.environ['WED'].replace(' ', '')
THU=os.environ['THU'].replace(' ', '')
FRI=os.environ['FRI'].replace(' ', '')
SAT=os.environ['SAT'].replace(' ', '')
HOL=os.environ['HOL'].replace(' ', '')