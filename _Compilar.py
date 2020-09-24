# -*- coding: cp1252 -*-
import py_compile
import shutil
from os import remove

shutil.copy('./LP_Config.xls','./Instalador')
print('LP_Config.xls copiado')

shutil.copy('./Considerações Técnicas N1 e N2 R05.pdf','./Instalador')
print('Considerações Técnicas copiado')

shutil.copy('./Padrao Planilha Supervisao_rev1P.xlsm','./Instalador')
print('Padrão Planilha Supervisão copiado')

py_compile.compile('FAS.pyw')
shutil.move('./__pycache__/FAS.cpython-34.pyc','./Instalador/FAS.pyw')
print('FAS.pywc compilado')

for arquivo in ['__init__.py',
                'base2lp.py',
                'cepel2lp.py',
                'Checar_LP.py',
                'func.py',
                'Gerar_LP.py',
                'Gerar_ONS.py',
                'gerarPlanilhaLP.py',
                'gerarPlanilhaONS.py',
                'LP.py',
                'LP_Comparar.py']:
    try:
        remove('./Instalador/lp_lib/__pycache__/{}.cpython-34.pyc'.format(arquivo.split('.')[0]))
    except:
        pass
    comp='./lp_lib/'+arquivo
    py_compile.compile(comp)
    shutil.move('./lp_lib/__pycache__/{}.cpython-34.pyc'.format(arquivo.split('.')[0]), './Instalador/lp_lib/{}c'.format(arquivo))
    print(arquivo,' compilado')

print() 
print()
input('\nPressione Enter...')
