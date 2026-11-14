"""Access-control tests (read-only; uses the already-recalculated /tmp copy)."""
from chama_core import Bot, XlsxBackend
bot=Bot(XlsxBackend('/tmp/test_chama.xlsx'),['254700000000'])
ADMIN='+254700000000'; JOHN='+254700000002'; MERCY='whatsapp:+254700000003'; STRANGER='254799999999'
def t(w,x,must=(),mustnot=()):
    o=bot.handle(w,x); print(f"\n[{w[-4:]}] {x}\n{o[:160]}")
    for m in must: assert m in o,(x,m,o)
    for m in mustnot: assert m not in o,(x,m,o)
t(JOHN,'bal',['John Kamau','14,300'],['Mercy'])
t(JOHN,'REG002',['Joined'],['Mercy'])
t(JOHN,'statement',['Statement for *John Kamau*'])
t(JOHN,'bal john',['John Kamau'])
t(JOHN,'REG003',['only show your own'],['Mercy'])
t(JOHN,'REG999',['only show your own'])
t(JOHN,'bal mercy',['only show your own'],['Mercy Achieng'])
t(JOHN,'bal m',['only show your own'],['Mercy','Faith','Samuel','Esther','David'])
t(JOHN,'statement REG004',['only show your own'],['Peter'])
t(JOHN,'arrears',['only show your own'],['Mercy'])
t(JOHN,'summary',['only show your own']); t(JOHN,'overdue',['only show your own'])
t(JOHN,'pay john 100',['only show your own']); t(JOHN,'loan john 5000 3',['only show your own'])
t(JOHN,'undo',['only show your own'])
t(MERCY,'REG003',['Mercy']); t(MERCY,'bal',['Mercy'],['John'])
t(STRANGER,'bal',['not registered']); t(STRANGER,'REG002',['not registered'],['John'])
t(STRANGER,'pay john 100',['not registered'])
t('',   'bal',['not registered'])
t(ADMIN,'REG003',['Mercy']); t(ADMIN,'bal john',['John']); t(ADMIN,'arrears',['Members owing'])
t(ADMIN,'pay john 2000',['Reply *YES*'])
print("\nACCESS TESTS PASSED")
o=bot.handle(ADMIN,'members'); print(o); assert 'REG001' in o and 'REG008' in o and 'Members (8)' in o
o=bot.handle(ADMIN,'0700000003'); assert 'Mercy Achieng' in o and 'Joined' in o
o=bot.handle(ADMIN,'+254700000004'); assert 'Peter Otieno' in o
o=bot.handle(ADMIN,'0799999999'); assert 'No member' in o
o=bot.handle(JOHN,'members'); assert 'only show your own' in o and 'Mercy' not in o
o=bot.handle(JOHN,'0700000003'); assert 'only show your own' in o and 'Mercy' not in o
print("ADMIN LOOKUP TESTS PASSED")
