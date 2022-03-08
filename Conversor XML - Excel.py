from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import xml.etree.ElementTree as ET
import pandas as pd
import os
import threading
import base64


icon = \
""" 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
Y6/O8f9sOrZxGj+siwAAAABJRU5ErkJggg==
"""

icondata = base64.b64decode(icon)
tempFile = "icon.ico"
iconfile = open(tempFile, "wb")
iconfile.write(icondata)
iconfile.close()

tkWindow = Tk()  
tkWindow.geometry('400x150')  
tkWindow.title('Conversor XML para Excel')
tkWindow.eval('tk::PlaceWindow . center')
tkWindow.wm_iconbitmap(tempFile)


def xmlnfe():

    messagebox.showinfo('Processo', 'Selecione a pasta com os XML')
    path = filedialog.askdirectory()

    messagebox.showinfo('Processo', 'Selecione onde deseja salvar o Relatório em Excel.')
    savefile = filedialog.asksaveasfilename()

    linha = 0
    linhasid = {'id': range(0, 20000)}

    numeronfe = ""
    data = ""
    cnpjemit = ""
    df = pd.DataFrame(linhasid)

    for filename in os.listdir(path):
        if not filename.endswith('.xml'): continue
        fullname = os.path.join(path, filename)
        parser = ET.XMLParser(encoding='iso-8859-5')
        tree = ET.parse(fullname, parser=parser)

        doc = tree.getroot()
        nodefind = doc.find('{http://www.portalfiscal.inf.br/nfe}NFe/{http://www.portalfiscal.inf.br/nfe}infNFe/{http://www.portalfiscal.inf.br/nfe}det')

        
        infCpl = ""
        
        for ide in doc.iter('{http://www.portalfiscal.inf.br/nfe}nNF'):
            
            numeronfe = ide.text
        
        for ide in doc.iter('{http://www.portalfiscal.inf.br/nfe}dhEmi'):
            
            data = ide.text

        for emit in doc.iter('{http://www.portalfiscal.inf.br/nfe}emit'):
        
            for CNPJ in emit.iter('{http://www.portalfiscal.inf.br/nfe}CNPJ'):
                
                cnpjemit = CNPJ.text
                
        for vNF in doc.iter('{http://www.portalfiscal.inf.br/nfe}vNF'):
            
            vnf = vNF.text
            

            for det in doc.iter ('{http://www.portalfiscal.inf.br/nfe}det'):
                
                nprod = det.attrib
                
                for xProd in det.iter('{http://www.portalfiscal.inf.br/nfe}xProd'):
                    
                    codprod = xProd.text

                    # VARIAVEIS PRODUTOS
                    #ICMS SN
                    csosn = ""
                    pcredsn = 0
                    icmssn = 0
                    orig = ""
                    #ICMS NORMAL
                    csticms = ""
                    vbcicms = 0
                    picms = 0
                    icmsnorm = 0
                    #ICMS ST
                    vbcst = 0
                    pst = 0
                    vst = 0
                    # FCP
                    pfcp = 0
                    vfcp = 0
                    # ICMS EFET
                    vbcefet = 0
                    pefet = 0
                    vefet = 0

                    # ICMS RET
                    vbcicmsret = 0
                    picmsret = 0
                    vICMSSubstituto = 0
                    vICMSSTRet = 0
                    
                    #IPI
                    cstipi = ""
                    vbcipi = 0
                    pipi = 0
                    vipi = 0

                    # PIS
                    cstpis = ""
                    vbcpis = 0
                    ppis = 0
                    vpis = 0
                    
                    # COFINS
                    cstcofins = ""
                    vbccofins = 0
                    pcofins = 0
                    vcofins = 0

                    # ICMS DIFAL
                    vBCUFDest = 0
                    pICMSUFDest = 0
                    vICMSUFDest = 0

                    # FRETES - DESCONTOS - OUTRAS
                    vFrete = 0
                    vDesc = 0
                    vOutro = 0

                    for NCM in det.iter('{http://www.portalfiscal.inf.br/nfe}NCM'):
                        
                        ncm = NCM.text
                        
                    for CFOP in det.iter('{http://www.portalfiscal.inf.br/nfe}CFOP'):
                        
                        cfop = CFOP.text

                    for vProd in det.iter('{http://www.portalfiscal.inf.br/nfe}vProd'):
                        vProd = vProd.text

                    for vFrete in det.iter('{http://www.portalfiscal.inf.br/nfe}vFrete'):
                        vFrete = vFrete.text

                    for vDesc in det.iter('{http://www.portalfiscal.inf.br/nfe}vDesc'):
                        vDesc = vDesc.text

                    for vOutro in det.iter('{http://www.portalfiscal.inf.br/nfe}vOutro'):
                        vOutro = vOutro.text

                    for ICMS in det.iter ('{http://www.portalfiscal.inf.br/nfe}ICMS'):

                        #ICMS SN
                        for CSOSN in ICMS.iter('{http://www.portalfiscal.inf.br/nfe}CSOSN'):
                            
                            csosn = CSOSN.text
                            
                        for pCredSN in ICMS.iter('{http://www.portalfiscal.inf.br/nfe}pCredSN'):
                            
                            pcredsn = pCredSN.text
                        
                        for vCredICMSSN in ICMS.iter('{http://www.portalfiscal.inf.br/nfe}vCredICMSSN'):
                            
                            icmssn = vCredICMSSN.text
                            
                        for orig in ICMS.iter('{http://www.portalfiscal.inf.br/nfe}orig'):
                            
                            orig = orig.text
                            
                        # ICMS NORMAL
                        for CST in ICMS.iter ('{http://www.portalfiscal.inf.br/nfe}CST'):
                            
                            csticms = CST.text
                            
                        for vBC in ICMS.iter ('{http://www.portalfiscal.inf.br/nfe}vBC'):
                            
                            vbcicms = vBC.text
                            
                        for pICMS in ICMS.iter ('{http://www.portalfiscal.inf.br/nfe}pICMS'):
                            
                            picms = pICMS.text
                            
                        for vICMS in ICMS.iter ('{http://www.portalfiscal.inf.br/nfe}vICMS'):
                            
                            icmsnorm = vICMS.text

                        # ICMS ST
                        for vBCST in ICMS.iter ('{http://www.portalfiscal.inf.br/nfe}vBCST'):
                            
                            vbcst = vBCST.text

                        for pICMSST in ICMS.iter ('{http://www.portalfiscal.inf.br/nfe}pICMSST'):
                            
                            pst = pICMSST.text
                            
                        for vICMSST in ICMS.iter ('{http://www.portalfiscal.inf.br/nfe}vICMSST'):
                            
                            vst = vICMSST.text

                        # ICMS EFET    
                        for vBCEfet in ICMS.iter ('{http://www.portalfiscal.inf.br/nfe}vBCEfet'):
                            
                            vbcefet = vBCEfet.text
                            
                        for pICMSEfet in ICMS.iter ('{http://www.portalfiscal.inf.br/nfe}pICMSEfet'):
                            
                            pefet = pICMSEfet.text
                            
                        for vICMSEfet in ICMS.iter ('{http://www.portalfiscal.inf.br/nfe}vICMSEfet'):
                            
                            vefet = vICMSEfet.text

                        # ICMS RET    
                        for vBCSTRet in ICMS.iter ('{http://www.portalfiscal.inf.br/nfe}vBCSTRet'):
                            
                            vbcicmsret = vBCSTRet.text
                            
                        for pST in ICMS.iter ('{http://www.portalfiscal.inf.br/nfe}pST'):
                            
                            picmsret = pST.text
                            
                        for vICMSSubstituto in ICMS.iter ('{http://www.portalfiscal.inf.br/nfe}vICMSSubstituto'):
                            
                            vICMSSubstituto = vICMSSubstituto.text

                        for vICMSSTRet in ICMS.iter ('{http://www.portalfiscal.inf.br/nfe}vICMSSTRet'):
                            
                            vICMSSTRet = vICMSSTRet.text

                        # FCP
                        for pFCP in ICMS.iter ('{http://www.portalfiscal.inf.br/nfe}pFCP'):
                            
                            pfcp = pFCP.text
                            
                        for vFCP in ICMS.iter ('{http://www.portalfiscal.inf.br/nfe}vFCP'):
                            
                            vfcp = vFCP.text
                            
                    for IPI in det.iter ('{http://www.portalfiscal.inf.br/nfe}IPI'):

                        for CST in IPI.iter ('{http://www.portalfiscal.inf.br/nfe}CST'):
                            
                            cstipi = CST.text
                            
                        for vBC in IPI.iter ('{http://www.portalfiscal.inf.br/nfe}vBC'):
                            
                            vbcipi = vBC.text
                            
                        for pIPI in IPI.iter ('{http://www.portalfiscal.inf.br/nfe}pIPI'):
                            
                            pipi = pIPI.text
                            
                        for vIPI in IPI.iter ('{http://www.portalfiscal.inf.br/nfe}vIPI'):
                            
                            vipi = vIPI.text

                    for PIS in det.iter ('{http://www.portalfiscal.inf.br/nfe}PIS'):

                        for CST in PIS.iter ('{http://www.portalfiscal.inf.br/nfe}CST'):
                            
                            cstpis = CST.text
                            
                        for vBC in PIS.iter ('{http://www.portalfiscal.inf.br/nfe}vBC'):
                            
                            vbcpis = vBC.text
                            
                        for pPIS in PIS.iter ('{http://www.portalfiscal.inf.br/nfe}pPIS'):
                            
                            ppis = pPIS.text
                            
                        for vPIS in PIS.iter ('{http://www.portalfiscal.inf.br/nfe}vPIS'):
                            
                            vpis = vPIS.text

                    for COFINS in det.iter ('{http://www.portalfiscal.inf.br/nfe}COFINS'):

                        for CST in COFINS.iter ('{http://www.portalfiscal.inf.br/nfe}CST'):
                            
                            cstcofins = CST.text
                            
                        for vBC in COFINS.iter ('{http://www.portalfiscal.inf.br/nfe}vBC'):
                            
                            vbccofins = vBC.text
                            
                        for pCOFINS in COFINS.iter ('{http://www.portalfiscal.inf.br/nfe}pCOFINS'):
                            
                            pcofins = pCOFINS.text
                            
                        for vCOFINS in COFINS.iter ('{http://www.portalfiscal.inf.br/nfe}vCOFINS'):
                            
                            vcofins = vCOFINS.text

                    for ICMSUFDest in det.iter ('{http://www.portalfiscal.inf.br/nfe}ICMSUFDest'):

                            
                        for vBCUFDest in ICMSUFDest.iter ('{http://www.portalfiscal.inf.br/nfe}vBCUFDest'):
                            
                            vBCUFDest = vBCUFDest.text
                            
                        for pICMSUFDest in ICMSUFDest.iter ('{http://www.portalfiscal.inf.br/nfe}pICMSUFDest'):
                            
                            pICMSUFDest = pICMSUFDest.text
                            
                        for vICMSUFDest in ICMSUFDest.iter ('{http://www.portalfiscal.inf.br/nfe}vICMSUFDest'):
                            
                            vICMSUFDest = vICMSUFDest.text

                    #for infCpl in doc.iter ('{http://www.portalfiscal.inf.br/nfe}infCpl'):
                        #infCpl = infCpl.text
                            
                    

                    df.loc[df['id'] == linha , 'NF'] = (numeronfe)
                    df.loc[df['id'] == linha , 'CNPJ'] = str(cnpjemit)
                    df.loc[df['id'] == linha , 'DATA'] = (data)
                    df.loc[df['id'] == linha , 'VALOR'] = float(vnf)
                    df.loc[df['id'] == linha , 'PRODUTO'] =str(nprod) + '-' + str(codprod)
                    df.loc[df['id'] == linha , 'NCM'] = str(ncm)
                    df.loc[df['id'] == linha , 'CFOP'] = (cfop)
                    df.loc[df['id'] == linha , 'ICMS CST'] = (csticms)
                    df.loc[df['id'] == linha , 'BC ICMS'] = float(vbcicms)
                    df.loc[df['id'] == linha , 'ALIQ ICMS'] = float(picms)
                    df.loc[df['id'] == linha , 'VALOR ICMS'] = float(icmsnorm)
                    df.loc[df['id'] == linha , 'ICMS CSOSN'] = (csosn)
                    df.loc[df['id'] == linha , 'ALIQ ICMS SN'] = float(pcredsn)
                    df.loc[df['id'] == linha , 'VALOR ICMS SN'] = float(icmssn)
                    df.loc[df['id'] == linha , 'BC ICMS ST'] = float(vbcst)
                    df.loc[df['id'] == linha , 'ALIQ ICMS ST'] = float(pst)
                    df.loc[df['id'] == linha , 'VALOR ICMS ST'] = float(vst)
                    df.loc[df['id'] == linha , 'ALIQ FCP'] = float(pfcp)
                    df.loc[df['id'] == linha , 'VALOR FCP'] = float(vfcp)
                    df.loc[df['id'] == linha , 'BC ICMS EFET'] = float(vbcefet)
                    df.loc[df['id'] == linha , 'ALIQ ICMS EFET'] = float(pefet)
                    df.loc[df['id'] == linha , 'VALOR ICMS EFET'] = float(vefet)
                    df.loc[df['id'] == linha , 'BC ICMS RET'] = float(vbcicmsret)
                    df.loc[df['id'] == linha , 'ALIQ ICMS RET'] = float(picmsret)
                    df.loc[df['id'] == linha , 'VALOR ICMS SUBSTITUTO'] = float(vICMSSubstituto)
                    df.loc[df['id'] == linha , 'VALOR ICMS RET'] = float(vICMSSTRet)
                    df.loc[df['id'] == linha , 'IPI CST'] = (cstipi)
                    df.loc[df['id'] == linha , 'BC IPI'] = float(vbcipi)
                    df.loc[df['id'] == linha , 'ALIQ IPI'] = float(pipi)
                    df.loc[df['id'] == linha , 'VALOR IPI'] = float(vipi)
                    df.loc[df['id'] == linha , 'VALOR PRODUTO'] = float(vProd) + float(vFrete) + float(vOutro) - float(vDesc)
                    df.loc[df['id'] == linha , 'BC PIS'] = float(vbcpis)
                    df.loc[df['id'] == linha , 'ALIQ PIS'] = float(ppis)
                    df.loc[df['id'] == linha , 'VALOR PIS'] = float(vpis)
                    df.loc[df['id'] == linha , 'COFINS CST'] = (cstcofins)
                    df.loc[df['id'] == linha , 'BC COFINS'] = float(vbccofins)
                    df.loc[df['id'] == linha , 'ALIQ COFINS'] = float(pcofins)
                    df.loc[df['id'] == linha , 'VALOR COFINS'] = float(vcofins)
                    df.loc[df['id'] == linha , 'BC DIFAL'] = float(vBCUFDest)
                    df.loc[df['id'] == linha , 'ALIQ DIFAL'] = float(pICMSUFDest)
                    df.loc[df['id'] == linha , 'VALOR DIFAL'] = float(vICMSUFDest)
                    df.loc[df['id'] == linha , 'ORIGEM ICMS'] = (orig)

                    linha = linha + 1

    df.to_excel(str(savefile)+".xlsx", index=False)

def xmlcte():

    messagebox.showinfo('Processo', 'Selecione a pasta com os XML')
    path = filedialog.askdirectory()

    messagebox.showinfo('Processo', 'Selecione onde deseja salvar o Relatório em Excel.')
    savefile = filedialog.asksaveasfilename ()

    linha = 0
    linhasid = {'id':range(0,20000)}

    numerocte = ""
    data = ""
    cnpjemit = ""
    df = pd.DataFrame(linhasid)


    for filename in os.listdir(path):
        if not filename.endswith('.xml'): continue
        fullname = os.path.join(path, filename)
        parser = ET.XMLParser(encoding='iso-8859-5')
        tree = ET.parse(fullname, parser=parser)

        doc = tree.getroot()
        nodefind = doc.find('{http://www.portalfiscal.inf.br/nfe}NFe/{http://www.portalfiscal.inf.br/nfe}infNFe/{http://www.portalfiscal.inf.br/nfe}det')
        
        
        infCpl = ""
        
        for ide in doc.iter('{http://www.portalfiscal.inf.br/cte}nCT'):
            
            numerocte = ide.text
        
        for ide in doc.iter('{http://www.portalfiscal.inf.br/cte}dhEmi'):
            
            data = ide.text

        for emit in doc.iter ('{http://www.portalfiscal.inf.br/cte}emit'):
        
            for CNPJ in emit.iter('{http://www.portalfiscal.inf.br/cte}CNPJ'):
                
                cnpjemit = CNPJ.text
                
        for vRec in doc.iter ('{http://www.portalfiscal.inf.br/cte}vRec'):
            
            vcte = vRec.text
            

                    # VARIAVEIS PRODUTOS
            chave = ""        
                    #ICMS NORMAL
            csticms = ""
            vbcicms = 0
            picms = 0
            icmsnorm = 0

                        
        for CFOP in doc.iter('{http://www.portalfiscal.inf.br/cte}CFOP'):
                        
                cfop = CFOP.text
                        
        for ICMS in doc.iter ('{http://www.portalfiscal.inf.br/cte}ICMS'):


                        # ICMS NORMAL
            for CST in ICMS.iter ('{http://www.portalfiscal.inf.br/cte}CST'):
                            
                            csticms = CST.text
                            
            for vBC in ICMS.iter ('{http://www.portalfiscal.inf.br/cte}vBC'):
                            
                            vbcicms = vBC.text
                            
            for pICMS in ICMS.iter ('{http://www.portalfiscal.inf.br/cte}pICMS'):
                            
                            picms = pICMS.text
                            
            for vICMS in ICMS.iter ('{http://www.portalfiscal.inf.br/cte}vICMS'):
                            
                            icmsnorm = vICMS.text
                        
        for infDoc in doc.iter ('{http://www.portalfiscal.inf.br/cte}infDoc'):
            for chave in infDoc.iter ('{http://www.portalfiscal.inf.br/cte}chave'):
                        
                chave = chave.text    
                            
                    
                  
                df.loc[df['id'] == linha , 'CTE'] = (numerocte)
                df.loc[df['id'] == linha , 'CNPJ'] = str(cnpjemit)
                df.loc[df['id'] == linha , 'DATA'] = (data)
                df.loc[df['id'] == linha , 'VALOR'] = float(vcte)    
                df.loc[df['id'] == linha , 'CFOP'] = (cfop)
                df.loc[df['id'] == linha , 'ICMS CST'] = (csticms)
                df.loc[df['id'] == linha , 'BC ICMS'] = float(vbcicms)
                df.loc[df['id'] == linha , 'ALIQ ICMS'] = float(picms)
                df.loc[df['id'] == linha , 'VALOR ICMS'] = float(icmsnorm)
                df.loc[df['id'] == linha , 'CHAVE'] = str(chave)
                    

                linha = linha + 1

    df.to_excel( str (savefile)+ ".xlsx" , index = False)

def xmlnfs():

    messagebox.showinfo('Processo', 'Selecione a pasta com os XML')
    path = filedialog.askdirectory()

    messagebox.showinfo('Processo', 'Selecione onde deseja salvar o Relatório em Excel.')
    savefile = filedialog.asksaveasfilename ()

    linha = 0
    linhasid = {'id':range(0,20000)}

    numerocte = ""
    data = ""
    cnpjemit = ""
    df = pd.DataFrame(linhasid)


    for filename in os.listdir(path):
        if not filename.endswith('.xml'): continue
        fullname = os.path.join(path, filename)
        parser = ET.XMLParser(encoding='iso-8859-5')
        tree = ET.parse(fullname, parser=parser)
        
        doc = tree.getroot()
        nodefind = doc.find('{http://www.portalfiscal.inf.br/nfe}NFe/{http://www.portalfiscal.inf.br/nfe}infNFe/{http://www.portalfiscal.inf.br/nfe}det')
        
        
        for NFS in doc.iter('NFS-e'):

            vserv = 0
            vtnf = 0
            vtliq = 0
            vretir = 0
            vretpis = 0
            vretcofins = 0
            vretcsll = 0
            vretinss = 0
            viss = 0
            vstiss = 0
            
            
            for Id in NFS.iter('nNFS-e'):
                
                numeronfs = Id.text
            
            for Id in NFS.iter('dEmi'):
                
                data = Id.text

            for prest in NFS.iter ('prest'):
            
                for CNPJ in prest.iter('CNPJ'):
                    
                    cnpjemit = CNPJ.text
                    
            for total in NFS.iter('total'):  

                

                for vtNF in total.iter ('vtNF'):
                    
                    vtnf = vtNF.text

                for vtLiq in total.iter ('vtLiq'):
                    
                    vtliq = vtLiq.text


            cLCServ = 0
            
            for det in NFS.iter ('det'):

                for vServ in det.iter ('vServ'):
                    
                    vserv = vServ.text

                for cLCServ in det.iter ('cLCServ'):
                    cLCServ = cLCServ.text

                for vRetIR in det.iter ('vRetIR'):
                    
                    vretir = vRetIR.text

                for vRetPISPASEP in det.iter ('vRetPISPASEP'):
                    
                    vretpis = vRetPISPASEP.text

                for vRetCOFINS in det.iter ('vRetCOFINS'):
                    
                    vretcofins = vRetCOFINS.text

                for vRetCSLL in det.iter ('vRetCSLL'):
                    
                    vretcsll = vRetCSLL.text

                for vRetINSS in det.iter ('vRetINSS'):
                    
                    vretinss = vRetINSS.text

                for vISS in det.iter ('vISS'):
                    
                    viss = vISS.text

                for vISSST in det.iter ('vISSST'):
                    
                    vstiss = vISSST.text
                    
                    
            
                    
                
          
                df.loc[df['id'] == linha , 'NFS'] = (numeronfs)
                df.loc[df['id'] == linha , 'CNPJ'] = str(cnpjemit)
                df.loc[df['id'] == linha , 'DATA'] = (data)
                df.loc[df['id'] == linha , 'CODIGO SERVIÇO'] = str(cLCServ)
                df.loc[df['id'] == linha , 'VALOR SERVIÇO'] = float(vserv)    
                df.loc[df['id'] == linha , 'VALOR NOTA'] = float(vtnf)
                df.loc[df['id'] == linha , 'VALOR LIQUIDO'] = float(vtliq)
                df.loc[df['id'] == linha , 'IR'] = float(vretir)
                df.loc[df['id'] == linha , 'PIS'] = float(vretpis)
                df.loc[df['id'] == linha , 'COFINS'] = float(vretcofins)
                df.loc[df['id'] == linha , 'CSLL'] = float(vretcsll)
                df.loc[df['id'] == linha , 'INSS'] = float(vretinss)
                df.loc[df['id'] == linha , 'ISS'] = float(viss)
                df.loc[df['id'] == linha , 'STISS'] = float(vstiss)
                    

                linha = linha + 1

    df.to_excel( str (savefile)+ ".xlsx" , index = False)

def processo(processo):
    loading_process = threading.Thread(target=processo)
    loading_process.start()
    loading()
    while loading_process.is_alive():
        tkWindow.config(cursor="circle")
        tkWindow.update()
    else:
        messagebox.showinfo('Processo', 'Processo concluido')
        tkWindow.config(cursor="")
        tkWindow.update()

def loading():
    messagebox.showinfo('Processando', 'Aguarde!')

def executarnfe():
    processo(xmlnfe)

def executarcte():
    processo(xmlcte)

def executarnfs():
    processo(xmlnfs)


button1 = Button(tkWindow,
	text = 'XML NFE',
	command = executarnfe)
button1.grid (row=1, column=0, columnspan=5, padx=100, ipadx=80)

button2 = Button(tkWindow,
	text = 'XML CTE',
	command = executarcte)
button2.grid (row=2, column=0, columnspan=5, padx=10, ipadx=81)

button3 = Button(tkWindow,
	text = 'XML NFS CXS',
	command = executarnfs)
button3.grid (row=3, column=0, columnspan=5, padx=10, ipadx=68)


tkWindow.mainloop()
