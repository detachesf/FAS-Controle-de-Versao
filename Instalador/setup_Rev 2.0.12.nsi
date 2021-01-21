# Auto-generated by EclipseNSIS Script Wizard
# 04/05/2012 13:33:14

Name "FAS"

# General Symbol Definitions
!define REGKEY "SOFTWARE\$(^Name)"
!define VERSION "2.0.12"
!define COMPANY "DETA (DEPARTAMENTO DE ENGENRARIA DE AUTOMA��O)"
!define URL "http://novaintranet.chesf.gov.br/DE/SET/DETA/SitePages/FAS.aspx"

# MUI Symbol Definitions
!define MUI_ICON .\lp_lib\chesf.ico
!define MUI_FINISHPAGE_NOAUTOCLOSE
!define MUI_STARTMENUPAGE_REGISTRY_ROOT HKLM
!define MUI_STARTMENUPAGE_NODISABLE
!define MUI_STARTMENUPAGE_REGISTRY_KEY ${REGKEY}
!define MUI_STARTMENUPAGE_REGISTRY_VALUENAME StartMenuGroup
!define MUI_STARTMENUPAGE_DEFAULTFOLDER "FAS"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall-colorful.ico"
!define MUI_UNFINISHPAGE_NOAUTOCLOSE

# Included files
!include Sections.nsh
!include MUI2.nsh

# Variables
Var StartMenuGroup

# P�ginas do instalador                 
!insertmacro MUI_PAGE_WELCOME                                   # p�gina de boas vindas
!insertmacro MUI_PAGE_LICENSE licdata.txt                       # p�gina de licensa
!insertmacro MUI_PAGE_COMPONENTS                                # p�gina de componentes a ser instalados
!insertmacro MUI_PAGE_DIRECTORY                                 # p�gina de escolha do diret�rio de instala��o
!insertmacro MUI_PAGE_STARTMENU Application $StartMenuGroup     # p�gina de escolha de atalho no menu iniciar
!insertmacro MUI_PAGE_INSTFILES                                 # p�gina de instala��o dos arquivos
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

# Installer languages
!insertmacro MUI_LANGUAGE PortugueseBR                          # linguagem de instalador

# Installer attributes
OutFile "FAS setup_${VERSION}.exe"                                  # nome do arquivo de instala��o aser gerado
InstallDir "C:\FAS_${VERSION}"                           
CRCCheck on                                                     # checagem de integridade do instalador antes de instalar programas
XPStyle on                                                      # estilo das janelas do instalador
ShowInstDetails show                                            # exibi��o dos detalhes durante instala��o dos arquivos
VIProductVersion 0.0.0.0
VIAddVersionKey ProductName "FAS"
VIAddVersionKey ProductVersion "${VERSION}"
VIAddVersionKey CompanyName "${COMPANY}"
VIAddVersionKey FileVersion "${VERSION}"
VIAddVersionKey FileDescription ""
VIAddVersionKey LegalCopyright ""
InstallDirRegKey HKLM "${REGKEY}" Path
ShowUninstDetails show

# Installer sections
# Executa se tiver selecionado Instalar FAS (verificado de alguma forma em SEC0000)
Section "!Instalar FAS" SEC0000
	SetShellVarContext all
    SetOutPath $INSTDIR
    SetOverwrite on
    File "FAS.pyw"
	CreateShortCut "$SMPROGRAMS\FAS\FAS.lnk" "$INSTDIR\FAS.pyw"
	CreateShortCut "$DESKTOP\FAS.lnk" "C:\Python34\pythonw.exe" "FAS.pyw" "$INSTDIR\lp_lib\chesf.ico"
    File LP_Config.xls
    File "Padrao Planilha Supervisao_rev1P.xlsm"
	File "Considera��es T�cnicas N1 e N2 R05.pdf"
    SetOutPath $INSTDIR\lp_lib
    File .\lp_lib\__init__.pyc
    File .\lp_lib\base2lp.pyc
    File .\lp_lib\cepel2lp.pyc	
	File .\lp_lib\chesf.ico
    File .\lp_lib\Checar_LP.pyc
	File .\lp_lib\LP_Comparar.pyc
    File .\lp_lib\func.pyc
    File .\lp_lib\Gerar_LP.pyc
    File .\lp_lib\Gerar_ONS.pyc
    File .\lp_lib\gerarPlanilhaLP.pyc
    File .\lp_lib\gerarPlanilhaONS.pyc
    File .\lp_lib\LP.pyc

    # Grava no registro informa��o de que instalou os componentes da se��o "Instalar FAS"
    WriteRegStr HKLM "${REGKEY}\Components" "Instalar FAS" 1
SectionEnd

# Executa se tiver selecionado instalar Python (verificado de alguma forma em SEC0001)
Section "Instalar Python 3.4.4, xlrd, xlsxwriter e openpyxl" SEC0001
    SetOutPath $INSTDIR
    SetOverwrite on
    
    # Extrai arquivo e espera execu��o do mesmo para continuar o script
    File python-3.4.4.msi
    ExecWait '"msiexec" /i "python-3.4.4.msi"'          # Executa o arquivo e espera finalizar
    Delete python-3.4.4.msi                             # Deleta arquivo
    
	# Copiar XlsxWriter e xlrd para Python 3.4.4
    SetOutPath "C:\Python34\Lib\site-packages"
    SetOverwrite off
    File "xlrd-0.9.2-py3.3.egg-info"
    SetOutPath "C:\Python34\Lib\site-packages\xlrd"
    File "xlrd\biffh.py"
    File "xlrd\book.py"
    File "xlrd\compdoc.py"
    SetOutPath "C:\Python34\Lib\site-packages\xlrd\doc"
    File "xlrd\doc\compdoc.html"
    File "xlrd\doc\xlrd.html"
    SetOutPath "C:\Python34\Lib\site-packages\xlrd\examples"
    File "xlrd\examples\namesdemo.xls"
    File "xlrd\examples\xlrdnameAPIdemo.py"
    SetOutPath "C:\Python34\Lib\site-packages\xlrd\examples\__pycache__"
    File "xlrd\examples\__pycache__\xlrdnameAPIdemo.cpython-33.pyc"
    SetOutPath "C:\Python34\Lib\site-packages\xlrd"
    File "xlrd\formatting.py"
    File "xlrd\formula.py"
    File "xlrd\info.py"
    File "xlrd\licences.py"
    File "xlrd\sheet.py"
    File "xlrd\timemachine.py"
    File "xlrd\xldate.py"
    File "xlrd\xlsx.py"
    File "xlrd\__init__.py"
  
    SetOutPath "C:\Python34\Lib\site-packages\"
    File "XlsxWriter-0.5.6-py3.4.egg-info"
    SetOutPath "C:\Python34\Lib\site-packages\xlsxwriter"
    File "xlsxwriter\app.py"
    File "xlsxwriter\chart.py"
    File "xlsxwriter\chartsheet.py"
    File "xlsxwriter\chart_area.py"
    File "xlsxwriter\chart_bar.py"
    File "xlsxwriter\chart_column.py"
    File "xlsxwriter\chart_line.py"
    File "xlsxwriter\chart_pie.py"
    File "xlsxwriter\chart_radar.py"
    File "xlsxwriter\chart_scatter.py"
    File "xlsxwriter\chart_stock.py"
    File "xlsxwriter\comments.py"
    File "xlsxwriter\compatibility.py"
    File "xlsxwriter\compat_collections.py"
    File "xlsxwriter\contenttypes.py"
    File "xlsxwriter\core.py"
    File "xlsxwriter\drawing.py"
    File "xlsxwriter\format.py"
    File "xlsxwriter\packager.py"
    File "xlsxwriter\relationships.py"
    File "xlsxwriter\sharedstrings.py"
    File "xlsxwriter\styles.py"
    File "xlsxwriter\table.py"
    File "xlsxwriter\theme.py"
    File "xlsxwriter\utility.py"
    File "xlsxwriter\vml.py"
    File "xlsxwriter\workbook.py"
    File "xlsxwriter\worksheet.py"
    File "xlsxwriter\xmlwriter.py"
    File "xlsxwriter\__init__.py"

    SetOutPath "C:\Python34\Lib\site-packages\openpyxl.egg-info"
    File "openpyxl-3.0.6\openpyxl.egg-info\dependency_links.txt"
    File "openpyxl-3.0.6\openpyxl.egg-info\PKG-INFO"   
    File "openpyxl-3.0.6\openpyxl.egg-info\requires.txt"   
    File "openpyxl-3.0.6\openpyxl.egg-info\SOURCES.txt"
    File "openpyxl-3.0.6\openpyxl.egg-info\top_level.txt"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl"
    File "openpyxl-3.0.6\openpyxl\__init__.py"
    File "openpyxl-3.0.6\openpyxl\_constants.py"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\cell"
    File "openpyxl-3.0.6\openpyxl\cell\__init__.py"
    File "openpyxl-3.0.6\openpyxl\cell\_writer.py"
    File "openpyxl-3.0.6\openpyxl\cell\cell.py"
    File "openpyxl-3.0.6\openpyxl\cell\text.py"    
    File "openpyxl-3.0.6\openpyxl\cell\read_only.py"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\chart"
    File "openpyxl-3.0.6\openpyxl\chart\__init__.py"
    File "openpyxl-3.0.6\openpyxl\chart\_3d.py"
    File "openpyxl-3.0.6\openpyxl\chart\_chart.py"
    File "openpyxl-3.0.6\openpyxl\chart\area_chart.py"
    File "openpyxl-3.0.6\openpyxl\chart\axis.py"
    File "openpyxl-3.0.6\openpyxl\chart\bar_chart.py"
    File "openpyxl-3.0.6\openpyxl\chart\bubble_chart.py"
    File "openpyxl-3.0.6\openpyxl\chart\chartspace.py"
    File "openpyxl-3.0.6\openpyxl\chart\data_source.py"
    File "openpyxl-3.0.6\openpyxl\chart\descriptors.py"
    File "openpyxl-3.0.6\openpyxl\chart\error_bar.py"
    File "openpyxl-3.0.6\openpyxl\chart\label.py"
    File "openpyxl-3.0.6\openpyxl\chart\layout.py"
    File "openpyxl-3.0.6\openpyxl\chart\legend.py"
    File "openpyxl-3.0.6\openpyxl\chart\line_chart.py"
    File "openpyxl-3.0.6\openpyxl\chart\marker.py"
    File "openpyxl-3.0.6\openpyxl\chart\picture.py"
    File "openpyxl-3.0.6\openpyxl\chart\pie_chart.py"
    File "openpyxl-3.0.6\openpyxl\chart\pivot.py"
    File "openpyxl-3.0.6\openpyxl\chart\plotarea.py"
    File "openpyxl-3.0.6\openpyxl\chart\print_settings.py"
    File "openpyxl-3.0.6\openpyxl\chart\radar_chart.py"
    File "openpyxl-3.0.6\openpyxl\chart\reader.py"
    File "openpyxl-3.0.6\openpyxl\chart\reference.py"
    File "openpyxl-3.0.6\openpyxl\chart\scatter_chart.py"
    File "openpyxl-3.0.6\openpyxl\chart\series.py"
    File "openpyxl-3.0.6\openpyxl\chart\series_factory.py"
    File "openpyxl-3.0.6\openpyxl\chart\shapes.py"
    File "openpyxl-3.0.6\openpyxl\chart\stock_chart.py"
    File "openpyxl-3.0.6\openpyxl\chart\surface_chart.py"
    File "openpyxl-3.0.6\openpyxl\chart\text.py"
    File "openpyxl-3.0.6\openpyxl\chart\title.py"
    File "openpyxl-3.0.6\openpyxl\chart\trendline.py"
    File "openpyxl-3.0.6\openpyxl\chart\updown_bars.py"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\chartsheet"
    File "openpyxl-3.0.6\openpyxl\chartsheet\__init__.py"
    File "openpyxl-3.0.6\openpyxl\chartsheet\chartsheet.py"
    File "openpyxl-3.0.6\openpyxl\chartsheet\custom.py"
    File "openpyxl-3.0.6\openpyxl\chartsheet\properties.py"
    File "openpyxl-3.0.6\openpyxl\chartsheet\protection.py"
    File "openpyxl-3.0.6\openpyxl\chartsheet\publish.py"
    File "openpyxl-3.0.6\openpyxl\chartsheet\relation.py"
    File "openpyxl-3.0.6\openpyxl\chartsheet\views.py"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\comments"
    File "openpyxl-3.0.6\openpyxl\comments\__init__.py"
    File "openpyxl-3.0.6\openpyxl\comments\author.py"
    File "openpyxl-3.0.6\openpyxl\comments\comment_sheet.py"
    File "openpyxl-3.0.6\openpyxl\comments\comments.py"
    File "openpyxl-3.0.6\openpyxl\comments\shape_writer.py"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\compat"
    File "openpyxl-3.0.6\openpyxl\compat\__init__.py"
    File "openpyxl-3.0.6\openpyxl\compat\abc.py"
    File "openpyxl-3.0.6\openpyxl\compat\numbers.py"
    File "openpyxl-3.0.6\openpyxl\compat\product.py"
    File "openpyxl-3.0.6\openpyxl\compat\singleton.py"    
    File "openpyxl-3.0.6\openpyxl\compat\strings.py"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\descriptors"
    File "openpyxl-3.0.6\openpyxl\descriptors\__init__.py"
    File "openpyxl-3.0.6\openpyxl\descriptors\base.py"
    File "openpyxl-3.0.6\openpyxl\descriptors\excel.py"
    File "openpyxl-3.0.6\openpyxl\descriptors\namespace.py"    
    File "openpyxl-3.0.6\openpyxl\descriptors\nested.py"
    File "openpyxl-3.0.6\openpyxl\descriptors\sequence.py"
    File "openpyxl-3.0.6\openpyxl\descriptors\serialisable.py"
    File "openpyxl-3.0.6\openpyxl\descriptors\slots.py"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\drawing"
    File "openpyxl-3.0.6\openpyxl\drawing\__init__.py"
    File "openpyxl-3.0.6\openpyxl\drawing\colors.py"
    File "openpyxl-3.0.6\openpyxl\drawing\connector.py"    	
    File "openpyxl-3.0.6\openpyxl\drawing\drawing.py"
    File "openpyxl-3.0.6\openpyxl\drawing\effect.py"
    File "openpyxl-3.0.6\openpyxl\drawing\fill.py"
    File "openpyxl-3.0.6\openpyxl\drawing\geometry.py"    	
    File "openpyxl-3.0.6\openpyxl\drawing\graphic.py"
    File "openpyxl-3.0.6\openpyxl\drawing\image.py"
    File "openpyxl-3.0.6\openpyxl\drawing\line.py"
    File "openpyxl-3.0.6\openpyxl\drawing\picture.py"    	
    File "openpyxl-3.0.6\openpyxl\drawing\properties.py" 
    File "openpyxl-3.0.6\openpyxl\drawing\relation.py"
    File "openpyxl-3.0.6\openpyxl\drawing\spreadsheet_drawing.py"    	
    File "openpyxl-3.0.6\openpyxl\drawing\text.py"
    File "openpyxl-3.0.6\openpyxl\drawing\xdr.py"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\formatting"
    File "openpyxl-3.0.6\openpyxl\formatting\__init__.py"
    File "openpyxl-3.0.6\openpyxl\formatting\formatting.py"
    File "openpyxl-3.0.6\openpyxl\formatting\rule.py"     
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\formula"
    File "openpyxl-3.0.6\openpyxl\formula\__init__.py"
    File "openpyxl-3.0.6\openpyxl\formula\tokenizer.py"
    File "openpyxl-3.0.6\openpyxl\formula\translate.py"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\packaging"
    File "openpyxl-3.0.6\openpyxl\packaging\__init__.py"
    File "openpyxl-3.0.6\openpyxl\packaging\core.py"
    File "openpyxl-3.0.6\openpyxl\packaging\extended.py"    	
    File "openpyxl-3.0.6\openpyxl\packaging\interface.py"
    File "openpyxl-3.0.6\openpyxl\packaging\manifest.py"
    File "openpyxl-3.0.6\openpyxl\packaging\relationship.py"
    File "openpyxl-3.0.6\openpyxl\packaging\workbook.py" 
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\pivot"
    File "openpyxl-3.0.6\openpyxl\pivot\__init__.py"
    File "openpyxl-3.0.6\openpyxl\pivot\cache.py"
    File "openpyxl-3.0.6\openpyxl\pivot\fields.py"    	
    File "openpyxl-3.0.6\openpyxl\pivot\record.py"
    File "openpyxl-3.0.6\openpyxl\pivot\table.py"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\reader"
    File "openpyxl-3.0.6\openpyxl\reader\__init__.py"
    File "openpyxl-3.0.6\openpyxl\reader\drawings.py"
    File "openpyxl-3.0.6\openpyxl\reader\excel.py"
    File "openpyxl-3.0.6\openpyxl\reader\strings.py"    
    File "openpyxl-3.0.6\openpyxl\reader\workbook.py"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\styles"
    File "openpyxl-3.0.6\openpyxl\styles\__init__.py"
    File "openpyxl-3.0.6\openpyxl\styles\alignment.py"
    File "openpyxl-3.0.6\openpyxl\styles\borders.py"    	
    File "openpyxl-3.0.6\openpyxl\styles\builtins.py"
    File "openpyxl-3.0.6\openpyxl\styles\cell_style.py"
    File "openpyxl-3.0.6\openpyxl\styles\colors.py"
    File "openpyxl-3.0.6\openpyxl\styles\differential.py"    	
    File "openpyxl-3.0.6\openpyxl\styles\fills.py"
    File "openpyxl-3.0.6\openpyxl\styles\fonts.py"
    File "openpyxl-3.0.6\openpyxl\styles\named_styles.py"
    File "openpyxl-3.0.6\openpyxl\styles\numbers.py"    	
    File "openpyxl-3.0.6\openpyxl\styles\protection.py" 
    File "openpyxl-3.0.6\openpyxl\styles\proxy.py"
    File "openpyxl-3.0.6\openpyxl\styles\styleable.py"    	
    File "openpyxl-3.0.6\openpyxl\styles\stylesheet.py"
    File "openpyxl-3.0.6\openpyxl\styles\table.py"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\utils"
    File "openpyxl-3.0.6\openpyxl\utils\__init__.py"
    File "openpyxl-3.0.6\openpyxl\utils\_accel.py"
    File "openpyxl-3.0.6\openpyxl\utils\bound_dictionary.py"    	
    File "openpyxl-3.0.6\openpyxl\utils\cell.py"
    File "openpyxl-3.0.6\openpyxl\utils\dataframe.py"
    File "openpyxl-3.0.6\openpyxl\utils\datetime.py"    	
    File "openpyxl-3.0.6\openpyxl\utils\escape.py"
    File "openpyxl-3.0.6\openpyxl\utils\exceptions.py"
    File "openpyxl-3.0.6\openpyxl\utils\formulas.py"
    File "openpyxl-3.0.6\openpyxl\utils\indexed_list.py"    	
    File "openpyxl-3.0.6\openpyxl\utils\inference.py" 
    File "openpyxl-3.0.6\openpyxl\utils\protection.py"
    File "openpyxl-3.0.6\openpyxl\utils\units.py"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\xml"
    File "openpyxl-3.0.6\openpyxl\xml\__init__.py"
    File "openpyxl-3.0.6\openpyxl\xml\constants.py"
    File "openpyxl-3.0.6\openpyxl\xml\functions.py"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\workbook"
    File "openpyxl-3.0.6\openpyxl\workbook\__init__.py"
    File "openpyxl-3.0.6\openpyxl\workbook\_writer.py"
    File "openpyxl-3.0.6\openpyxl\workbook\child.py"
    File "openpyxl-3.0.6\openpyxl\workbook\defined_name.py"    
    File "openpyxl-3.0.6\openpyxl\workbook\external_reference.py"
    File "openpyxl-3.0.6\openpyxl\workbook\function_group.py"
    File "openpyxl-3.0.6\openpyxl\workbook\properties.py"
    File "openpyxl-3.0.6\openpyxl\workbook\protection.py"
    File "openpyxl-3.0.6\openpyxl\workbook\smart_tags.py"    
    File "openpyxl-3.0.6\openpyxl\workbook\views.py"
    File "openpyxl-3.0.6\openpyxl\workbook\web.py"    
    File "openpyxl-3.0.6\openpyxl\workbook\workbook.py"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\workbook\external_link"
    File "openpyxl-3.0.6\openpyxl\workbook\external_link\__init__.py"
    File "openpyxl-3.0.6\openpyxl\workbook\external_link\external.py"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\worksheet"
    File "openpyxl-3.0.6\openpyxl\worksheet\__init__.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\_read_only.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\_reader.py"    	
    File "openpyxl-3.0.6\openpyxl\worksheet\_write_only.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\_writer.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\cell_range.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\cell_watch.py"    	
    File "openpyxl-3.0.6\openpyxl\worksheet\controls.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\copier.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\custom.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\datavalidation.py"    	
    File "openpyxl-3.0.6\openpyxl\worksheet\dimensions.py" 
    File "openpyxl-3.0.6\openpyxl\worksheet\drawing.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\errors.py"    	
    File "openpyxl-3.0.6\openpyxl\worksheet\filters.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\header_footer.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\hyperlink.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\merge.py"    	
    File "openpyxl-3.0.6\openpyxl\worksheet\ole.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\copier.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\page.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\pagebreak.py"    	
    File "openpyxl-3.0.6\openpyxl\worksheet\picture.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\properties.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\protection.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\related.py"    	
    File "openpyxl-3.0.6\openpyxl\worksheet\scenario.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\smart_tag.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\table.py"
    File "openpyxl-3.0.6\openpyxl\worksheet\views.py"    	
    File "openpyxl-3.0.6\openpyxl\worksheet\worksheet.py"
    SetOutPath "C:\Python34\Lib\site-packages\openpyxl\writer"   
    File "openpyxl-3.0.6\openpyxl\writer\__init__.py"
    File "openpyxl-3.0.6\openpyxl\writer\excel.py"
    File "openpyxl-3.0.6\openpyxl\writer\theme.py"
    

    # Grava no registro informa��o de que instalou os componentes da se��o "Instalar Python 3.4.4"   
    WriteRegStr HKLM "${REGKEY}\Components" "Instalar Python 3.4.4, xlrd e xlsxwriter" 1    
SectionEnd

# Executado ap�s as se��es anteriores, "p�s instala��o"
Section -post SEC0002
	SetShellVarContext all
    WriteRegStr HKLM "${REGKEY}" Path $INSTDIR
    SetOutPath $INSTDIR
    WriteUninstaller $INSTDIR\uninstall.exe
    !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
    SetOutPath $SMPROGRAMS\$StartMenuGroup
    # Cria atalhos no menu iniciar 
    CreateShortcut "$SMPROGRAMS\$StartMenuGroup\Uninstall $(^Name).lnk" $INSTDIR\uninstall.exe
    # Altera diret�rio de trabalho inicial para os pr�ximos links, o campo "iniciar em" do link ter�  
    # o caminho do diret�rio de instala��o que foi escolhido para os scripts
    SetOutPath $INSTDIR         
    CreateShortcut "$SMPROGRAMS\$StartMenuGroup\FAS.lnk" "C:\Python34\pythonw.exe" "$INSTDIR\FAS.pyw"
    # Volta para o caminho anterior
    SetOutPath $SMPROGRAMS\$StartMenuGroup     
    
    !insertmacro MUI_STARTMENU_WRITE_END
    WriteRegStr HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$(^Name)" DisplayName "$(^Name)"
    WriteRegStr HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$(^Name)" DisplayVersion "${VERSION}"
    WriteRegStr HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$(^Name)" Publisher "${COMPANY}"
    WriteRegStr HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$(^Name)" DisplayIcon $INSTDIR\uninstall.exe
    WriteRegStr HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$(^Name)" UninstallString $INSTDIR\uninstall.exe
    WriteRegDWORD HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$(^Name)" NoModify 1
    WriteRegDWORD HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$(^Name)" NoRepair 1
SectionEnd

# Macro for selecting uninstaller sections
!macro SELECT_UNSECTION SECTION_NAME UNSECTION_ID
    Push $R0
    ReadRegStr $R0 HKLM "${REGKEY}\Components" "${SECTION_NAME}"
    StrCmp $R0 1 0 next${UNSECTION_ID}
    !insertmacro SelectSection "${UNSECTION_ID}"
    GoTo done${UNSECTION_ID}
next${UNSECTION_ID}:
    !insertmacro UnselectSection "${UNSECTION_ID}"
done${UNSECTION_ID}:
    Pop $R0
!macroend

# Uninstaller sections
Section /o "-un.Instalar Python 3.4.4" UNSEC0001
    DeleteRegValue HKLM "${REGKEY}\Components" "Instalar Python 3.4.4, xlrd e xlsxwriter"
SectionEnd

Section /o "-un.Instalar FAS" UNSEC0000
    # /REBOOTOK - se o arquivo n�o puder ser deletado (estiver em uso), ele � deletado
    # ap�s o reboot do computador
	SetShellVarContext all
    Delete /REBOOTOK $INSTDIR\lp_lib\gerarPlanilhaLP.pyc
    Delete /REBOOTOK $INSTDIR\lp_lib\gerarPlanilhaONS.pyc
    Delete /REBOOTOK $INSTDIR\lp_lib\base2lp.pyc
	Delete /REBOOTOK $INSTDIR\lp_lib\cepel2lp.pyc
    Delete /REBOOTOK $INSTDIR\lp_lib\func.pyc
    Delete /REBOOTOK $INSTDIR\lp_lib\__init__.pyc
    Delete /REBOOTOK $INSTDIR\lp_lib\LP.pyc
    Delete /REBOOTOK "$INSTDIR\Padrao Planilha Supervisao_rev1P.xlsm"
	Delete /REBOOTOK "$INSTDIR\Considera��es T�cnicas N1 e N2 R05.pdf"
    Delete /REBOOTOK $INSTDIR\LP_Config.xls
    Delete /REBOOTOK $INSTDIR\lp_lib\Checar_LP.pyc
	Delete /REBOOTOK $INSTDIR\lp_lib\LP_Comparar.pyc
    Delete /REBOOTOK $INSTDIR\lp_lib\Gerar_LP.pyc
	Delete /REBOOTOK $INSTDIR\lp_lib\Gerar_ONS.pyc
	Delete /REBOOTOK "$INSTDIR\FAS.pyw"
	Delete /REBOOTOK $INSTDIR\lp_lib\chesf.ico
	Delete /REBOOTOK $INSTDIR\lp_lib
	Delete /REBOOTOK "$SMPROGRAMS\$StartMenuGroup\Uninstall $(^Name).lnk"
    Delete /REBOOTOK "$SMPROGRAMS\$StartMenuGroup\FAS.lnk" 
	RmDir /REBOOTOK $SMPROGRAMS\$StartMenuGroup
    RmDir /REBOOTOK $INSTDIR
    DeleteRegValue HKLM "${REGKEY}\Components" "Instalar FAS"
	Delete "$DESKTOP\PyMEC1000.lnk"
SectionEnd

Section -un.post UNSEC0002
    DeleteRegKey HKLM "SOFTWARE\Microsoft\Windows\CurentVersion\Uninstall\$(^Name)"
    Delete /REBOOTOK $INSTDIR\uninstall.exe    
    DeleteRegValue HKLM "${REGKEY}" StartMenuGroup
    DeleteRegValue HKLM "${REGKEY}" Path
    DeleteRegKey /IfEmpty HKLM "${REGKEY}\Components"
    DeleteRegKey /IfEmpty HKLM "${REGKEY}"

SectionEnd

# Installer functions
Function .onInit
    InitPluginsDir
FunctionEnd

# Uninstaller functions
Function un.onInit
    ReadRegStr $INSTDIR HKLM "${REGKEY}" Path
    !insertmacro MUI_STARTMENU_GETFOLDER Application $StartMenuGroup
    !insertmacro SELECT_UNSECTION "Instalar FAS" ${UNSEC0000}
    !insertmacro SELECT_UNSECTION "Instalar Python 3.4.4, xlrd e XlsxWriter" ${UNSEC0001}
FunctionEnd

# Section Descriptions
!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
!insertmacro MUI_DESCRIPTION_TEXT ${SEC0000} "Instala FAS"
!insertmacro MUI_DESCRIPTION_TEXT ${SEC0001} "Instala Python 3.4.4 e bibliotecas"
!insertmacro MUI_FUNCTION_DESCRIPTION_END
