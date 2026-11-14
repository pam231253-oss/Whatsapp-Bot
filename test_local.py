"""Runs the bot against a COPY of the real workbook (formulas recalculated by LibreOffice)."""
import shutil, sys
from chama_core import Bot, XlsxBackend
shutil.copy('/mnt/user-data/outputs/Chama_Manager.xlsx','/tmp/test_chama.xlsx')
bot=Bot(XlsxBackend('/tmp/test_chama.xlsx'),['254700000000'])
ADMIN='+254700000000'; MEMBER='254700000002'
def say(who,txt):
    print(f"\n>>> {txt}"); out=bot.handle(who,txt); print(out); return out
say(ADMIN,'help')
o=say(ADMIN,'REG002'); assert 'Joined 01 Jan 2026' in o and '21,300' in o
o=say(ADMIN,'reg004'); assert 'Peter Otieno' in o
o=say('254700000003','REG002'); assert 'only' in o
o=say('254700000003','REG003'); assert 'Mercy' in o
say(ADMIN,'REG999')
o=say(ADMIN,'bal john'); assert '21,300' in o
say(ADMIN,'pay john 2000'); o=say(ADMIN,'yes'); assert '19,300' in o
o=say(ADMIN,'undo'); o=say(ADMIN,'bal john'); assert '21,300' in o
say(ADMIN,'pay john 4000 loan'); o=say(ADMIN,'yes')
say(ADMIN,'fine mercy 100 late contribution'); say(ADMIN,'yes')
say(ADMIN,'loan faith 5000 3'); o=say(ADMIN,'yes'); assert 'L005' in o
say(ADMIN,'pay grace 500 mgr'); say(ADMIN,'no')
say(ADMIN,'statement john')
say(ADMIN,'arrears'); say(ADMIN,'overdue'); say(ADMIN,'summary')
say(ADMIN,'pay sam')
say(ADMIN,'bal xyz')
say(MEMBER,'bal')
o=say(MEMBER,'pay john 100'); assert 'only' in o.lower()
o=say(MEMBER,'bal mercy'); assert 'own' in o
sms="TKB4X9L2QA Confirmed. Ksh2,000.00 received from JOHN KAMAU 254700000002 on 14/11/26 at 3:45 PM New Account balance is Ksh9,000.00."
say(ADMIN,sms); o=say(ADMIN,'yes')
print("\nALL ASSERTS PASSED")
