# coding: utf-8
import wx
import dataset
import xlsxwriter
from xlsxwriter.workbook import Workbook
import xlwt
import tagtool
import codecs
import csv
import csvkit
from stuf import stuf
# import the newly created GUI file
import JukuPlanner
import subprocess
import os
import sys
import glob
import funzioni
import calendar
import ConfigParser
import hashlib
from time import gmtime, strftime, sleep, localtime
from datetime import datetime
parser = ConfigParser.SafeConfigParser()
parser.read('./cfg.ini')
giorno=''
#Config = ConfigParser.SafeConfigParser()
#Config.read('./cfg.ini')
#pathdatabase = Config.get('Paths','databasepath')
#percorsoDatabase='sqlite:///'+pathdatabase
Freq = 1500  # Set Frequency To 2500 Hertz
Dur = 250  # Set Duration To 1000 ms == 1 second
db = dataset.connect('sqlite:///users.db', row_type=stuf)
dbins = dataset.connect('sqlite:///teacher.db', row_type=stuf)

#db = dataset.connect(percorsoDatabase, row_type=stuf)

tabella = db["users"]
tabellaTempo = db['timeTable']
tabellaGiorni = db['giorni']
tabellaCalcoli = db['calcoli']
tabellaIns = dbins['insegnanti']
tabellaTempoIns = dbins['timeTable']
tabellaDateIns = dbins['datePersonalizzate']
contaPrivate = 0

#settingsdb = dataset.connect('sqlite:///settings.db', row_type=stuf)
#tabellaSettings = settingsdb['settaggi']
colonna = 0
riga = 0
rigaSelezionata = 0
colonnaSelezionata = 0
rigaMaterie = 0
colonnaMaterie = 0
rigaMaterie1 = 0
colonnaMaterie1 = 0
idSelezionato = 0
idDatePersonalizzate = 0
idGiorni=0
idCalcoli = 0
stanza  = 0
datavecchia= ''
percorso = ''
materia = ''
materieArray=[]
materieTesto = []
switchmaterie=0
switchmaterieOriginale=0
mostraInfoStud = False
datiInfoStudente = ''
copia1 = ' '
coordinateCopia1 = []
copia2 = ' '
coordinateCopia2 = []
copia1m = ' '
coordinateCopia1m = []
copia2m = ' '
coordinateCopia2m = []
copia1Colore = ''
copia2Colore = ''
copia1Kojin = False
copia2Kojin = False
# importing * : to enable writing sin(13) instead of math.sin(13)先生
from math import *
class UTF8Recoder:
    """
    Iterator that reads an encoded stream and reencodes the input to UTF-8
    """
    def __init__(self, f, encoding):
        self.reader = codecs.getreader(encoding)(f)

    def __iter__(self):
        return self

    def next(self):
        return self.reader.next().encode("utf-8")
def unicode_csv_reader(unicode_csv_data, dialect=csv.excel, **kwargs):
    # csv.py doesn't do Unicode; encode temporarily as UTF-8:
    csv_reader = csv.reader(utf_8_encodermultiplo(unicode_csv_data),
                            dialect=dialect, **kwargs)
    dizionario = {}
    for row in csv_reader:

        # decode UTF-8 back to Unicode, cell by cell:
        yield [unicode(cell, 'utf-8') for cell in row]

def utf_8_encoder(unicode_csv_data):
    return unicode_csv_data.encode('utf-8')
def utf_8_encodermultiplo(unicode_csv_data):
    for line in unicode_csv_data:
        yield line.encode('utf-8')


class UnicodeReader:
    """
    A CSV reader which will iterate over lines in the CSV file "f",
    which is encoded in the given encoding.
    """

    def __init__(self, f, dialect=csv.excel, encoding="utf-8", **kwds):
        f = UTF8Recoder(f, encoding)
        self.reader = csv.reader(f, dialect=dialect, **kwds)

    def next(self):
        row = self.reader.next()
        return [unicode(s, "utf-8") for s in row]

    def __iter__(self):
        return self

# inherit from the MainFrame created in wxFowmBuilder and create CalcFrame
class CalcFrame(JukuPlanner.FramePrincipale):
    # constructor先生
    def __init__(self, parent):
        # initialize parent class
        Config = ConfigParser.ConfigParser()
        Config.read('./cfg.ini')
        colorestanza1 = Config.get('Colori stanze','colorestanza1')
        colorestanza2 = Config.get('Colori stanze', 'colorestanza2')
        colorestanza3 = Config.get('Colori stanze', 'colorestanza3')
        colorestanza4 = Config.get('Colori stanze', 'colorestanza4')
        percorsocsv = Config.get('Paths','csvpath')
        colore1 = funzioni.coonvertiStringaInColore(colorestanza1)
        colore2 = funzioni.coonvertiStringaInColore(colorestanza2)
        colore3 = funzioni.coonvertiStringaInColore(colorestanza3)
        colore4 = funzioni.coonvertiStringaInColore(colorestanza4)

        JukuPlanner.FramePrincipale.__init__(self, parent)
        insegnante = u"先生"
        studente = u"生徒"
        materia = u"科目"
        room1 = u'部屋1'
        room2 = u'部屋2'
        room3 = u'部屋3'
        room4 = u'部屋4'

        global datavecchia
        datavecchia = str(self.calendario.Date)
        self.griglia.SetColLabelValue(0, "9:10 - 10:20")
        self.griglia.SetColLabelValue(1, "10:30 - 11:40")
        self.griglia.SetColLabelValue(2, "11:50 - 13:00")
        self.griglia.SetColLabelValue(3, "13:40 - 14:50")
        self.griglia.SetColLabelValue(4, "15:00 - 16:10")
        self.griglia.SetColLabelValue(5, "16:40 - 17:50")
        self.griglia.SetColLabelValue(6, "18:00 - 19:10")
        self.griglia.SetColLabelValue(7, "19:20 - 20:30")
        self.griglia.SetColLabelValue(8, "20:40 - 21:50")
        self.griglia.SetRowLabelValue(0, insegnante)
        self.griglia.SetRowLabelValue(1, studente)
        self.griglia.SetRowLabelValue(2, studente)
        self.griglia.SetRowLabelValue(3, studente)
        self.griglia.SetRowLabelValue(4, studente)
        self.griglia.SetRowLabelValue(5, insegnante)
        self.griglia.SetRowLabelValue(6, studente)
        self.griglia.SetRowLabelValue(7, studente)
        self.griglia.SetRowLabelValue(8, studente)
        self.griglia.SetRowLabelValue(9, studente)
        self.griglia.SetRowLabelValue(10, insegnante)
        self.griglia.SetRowLabelValue(11, studente)
        self.griglia.SetRowLabelValue(12, studente)
        self.griglia.SetRowLabelValue(13, studente)
        self.griglia.SetRowLabelValue(14, studente)
        self.griglia.SetRowLabelValue(15, insegnante)
        self.griglia.SetRowLabelValue(16, studente)
        self.griglia.SetRowLabelValue(17, studente)
        self.griglia.SetRowLabelValue(18, studente)
        self.griglia.SetRowLabelValue(19, studente)
        for i in range(0, 9, 1):
            #self.griglia.SetCellBackgroundColour(0, i, wx.GREEN)
            self.griglia.SetCellBackgroundColour(0, i,wx.Colour(int(colore1[0]), int(colore1[1]), int(colore1[2]), int(colore1[3])))
            self.griglia.SetCellBackgroundColour(5, i, wx.Colour(int(colore2[0]), int(colore2[1]), int(colore2[2]), int(colore2[3])))
            self.griglia.SetCellBackgroundColour(10, i,wx.Colour(int(colore3[0]), int(colore3[1]), int(colore3[2]), int(colore3[3])))
            self.griglia.SetCellBackgroundColour(15, i, wx.Colour(int(colore4[0]), int(colore4[1]), int(colore4[2]), int(colore4[3])))
        self.griglia.SetColSize(0, 100)
        self.griglia.SetColSize(1, 100)
        self.griglia.SetColSize(2, 100)
        self.griglia.SetColSize(3, 100)
        self.griglia.SetColSize(4, 100)
        self.griglia.SetColSize(5, 100)
        self.griglia.SetColSize(6, 100)
        self.griglia.SetColSize(7, 100)
        self.griglia.SetColSize(8, 100)


        popolaInsegnanti = tabella.find(teacher='1')
        popolaStudenti = tabella.find(student='1',)
        listaMaterie = [u'国語', u'英語', u'数学', u'理科', u'社会', u'特別']
        #for i in popolaStudenti:
        #    self.listaStudenti.Append(i.name)
        for i in popolaInsegnanti:
            self.listaInsegnanti.Append(i.name)
        for i in listaMaterie:
            self.listaMaterie.Append(i)
        nomeFile = str(self.calendario.Date)

        nomeFile = nomeFile.replace('/', '-')
        nomeFile = nomeFile.replace(' 00:00:00', '')
        anno = '20' + nomeFile[-2:]
        global percorso
        percorso = percorsocsv+'/' + anno + '/' + nomeFile[:2] + '/' + nomeFile + '.csv'

        if not os.path.exists(os.path.dirname(percorso)):
            try:
                os.makedirs(os.path.dirname(percorso))
            except OSError as exc:  # Guard against race condition
                pass
        print percorso
        controllaPercorso = os.path.exists(percorso)

        if controllaPercorso == True:
            with open(percorso, 'rb') as f:
                reader = csv.DictReader(f)
                contarighe = 0
                converti = csvkit.unicsv.UnicodeCSVDictReader(f=f, encoding='utf-8')

                for i in converti:
                    print i, 'i', type(i)
                    self.griglia.SetCellValue(contarighe, 0, i['9:10 - 10:20'])
                    self.griglia.SetCellValue(contarighe, 1, i['10:30 - 11:40'])
                    self.griglia.SetCellValue(contarighe, 2, i['11:50 - 13:00'])
                    self.griglia.SetCellValue(contarighe, 3, i['13:40 - 14:50'])
                    self.griglia.SetCellValue(contarighe, 4, i['15:00 - 16:10'])
                    self.griglia.SetCellValue(contarighe, 5, i['16:40 - 17:50'])
                    self.griglia.SetCellValue(contarighe, 6, i['18:00 - 19:10'])
                    self.griglia.SetCellValue(contarighe, 7, i['19:20 - 20:30'])
                    self.griglia.SetCellValue(contarighe, 8, i['20:40 - 21:50'])
                    contarighe = contarighe + 1
    def inserimentoAutomatico( self, event ):
        lista = []
        nonInseriti = []
        dizionario = dict()

        for i in self.studentiDelGiorno.Items:

            risultati = self.preparativiInserimentoAutomatico(i)
            #lista.append(risultati)
            dizionario[i]=risultati

        print dizionario, 'risultati'
        for dizio in dizionario:
            soloIndividuali = False
            studente = tabella.find_one(name=dizio, student=1)
            if studente.individual == True and studente.shared == False:
                soloIndividuali = True
            if soloIndividuali == True:
                print dizio, 'supporta solo lezioni singole'
                print 'cerco stanze disponibili'
                print dizionario[dizio]
                contaore = 0

                for diz in dizionario[dizio]:
                    for i in range(0, self.griglia.NumberRows):
                        if dizio in self.griglia.GetCellValue(i,contaore):
                            inserito= True

                            break
                        else:
                            #print 'Elemento da inserire', dizio
                            inserito = False

                    if diz != u'':
                        if self.griglia.GetCellValue(1,contaore) == '' and inserito == False:
                            self.griglia.SetCellValue(1, contaore,
                                                      '(K)' + unicode(dizio) + u' ' + u'(' + unicode(diz.strip('K')) + u')')
                            self.griglia.SetCellValue(2,contaore,'(K)')
                            self.griglia.SetCellValue(3, contaore,
                                                      '(K)')
                            self.griglia.SetCellValue(4, contaore,
                                                      '(K)' )

                            #inserito == True
                        elif self.griglia.GetCellValue(6  ,contaore) == '' and inserito == False:
                            self.griglia.SetCellValue(6,contaore,'(K)'+unicode(dizio)+u' '+u'('+unicode(diz.strip('K'))+u')')
                            self.griglia.SetCellValue(7, contaore, '(K)')
                            self.griglia.SetCellValue(8, contaore,
                                                      '(K)')
                            self.griglia.SetCellValue(9, contaore,
                                                      '(K)')
                            #inserito == True
                        elif self.griglia.GetCellValue(11  ,contaore) == '' and inserito == False:
                            self.griglia.SetCellValue(11,contaore,'(K)'+unicode(dizio)+u' '+u'('+unicode(diz.strip('K'))+u')')
                            self.griglia.SetCellValue(12, contaore, '(K)')
                            self.griglia.SetCellValue(13, contaore,
                                                      '(K)')
                            self.griglia.SetCellValue(14, contaore,
                                                      '(K)')
                            #inserito == True
                        elif self.griglia.GetCellValue(16  ,contaore) == '' and inserito == False:
                            self.griglia.SetCellValue(16,contaore,'(K)'+unicode(dizio)+u' '+u'('+unicode(diz.strip('K'))+u')')
                            self.griglia.SetCellValue(17, contaore, '(K)')
                            self.griglia.SetCellValue(18, contaore,
                                                      '(K)')
                            self.griglia.SetCellValue(19, contaore,
                                                      '(K)')
                        elif self.griglia.GetCellValue(16  ,contaore) != '' and inserito == False:
                            if contaore == 0:
                                nonInseriti.append(dizio + ' ' + '9:10')
                            elif contaore == 1:
                                nonInseriti.append(dizio + ' ' + '10:30')
                            elif contaore == 2:
                                nonInseriti.append(dizio + ' ' + '11:50')
                            elif contaore == 3:
                                nonInseriti.append(dizio + ' ' + '13:40')
                            elif contaore == 4:
                                nonInseriti.append(dizio + ' ' + '15:00')
                            elif contaore == 5:
                                nonInseriti.append(dizio + ' ' + '16:40')
                            elif contaore == 6:
                                nonInseriti.append(dizio + ' ' + '18:00')
                            elif contaore == 7:
                                nonInseriti.append(dizio + ' ' + '19:20')
                            elif contaore == 8:
                                nonInseriti.append(dizio + ' ' + '20:40')


                            #inserito == True
                    contaore = contaore + 1
            if soloIndividuali == False:
                print dizio, 'supporta lezioni di gruppo'
                print 'cerco stanze disponibili'
                print dizionario[dizio]
                contaore = 0

                for diz in dizionario[dizio]:
                    for i in range(0, self.griglia.NumberRows):
                        if dizio in self.griglia.GetCellValue(i,contaore):
                            inserito= True
                            break
                        else:
                            inserito = False
                    if u'K' in diz:
                        if self.griglia.GetCellValue(1,contaore) == '' and inserito == False:
                            self.griglia.SetCellValue(1, contaore,
                                                      unicode(dizio) + u' ' + u'(' + unicode(diz) + u')')
                            self.griglia.SetCellValue(2,contaore,'(K)')
                            self.griglia.SetCellValue(3, contaore,
                                                      '(K)')
                            self.griglia.SetCellValue(4, contaore,
                                                      '(K)' )

                            #inserito == True
                        elif self.griglia.GetCellValue(6  ,contaore) == '' and inserito == False:
                            self.griglia.SetCellValue(6,contaore,unicode(dizio)+u' '+u'('+unicode(diz)+u')')
                            self.griglia.SetCellValue(7, contaore, '(K)')
                            self.griglia.SetCellValue(8, contaore,
                                                      '(K)')
                            self.griglia.SetCellValue(9, contaore,
                                                      '(K)')
                            #inserito == True
                        elif self.griglia.GetCellValue(11  ,contaore) == '' and inserito == False:
                            self.griglia.SetCellValue(11,contaore,unicode(dizio)+u' '+u'('+unicode(diz)+u')')
                            self.griglia.SetCellValue(12, contaore, '(K)')
                            self.griglia.SetCellValue(13, contaore,
                                                      '(K)')
                            self.griglia.SetCellValue(14, contaore,
                                                      '(K)')
                            #inserito == True
                        elif self.griglia.GetCellValue(16  ,contaore) == '' and inserito == False:
                            self.griglia.SetCellValue(16,contaore,unicode(dizio)+u' '+u'('+unicode(diz)+u')')
                            self.griglia.SetCellValue(17, contaore, '(K)')
                            self.griglia.SetCellValue(18, contaore,
                                                      '(K)')
                            self.griglia.SetCellValue(19, contaore,
                                                      '(K)')
                        elif self.griglia.GetCellValue(16  ,contaore) != '' and inserito == False:
                            if contaore == 0:
                                nonInseriti.append(dizio + ' ' + '9:10')
                            elif contaore == 1:
                                nonInseriti.append(dizio + ' ' + '10:30')
                            elif contaore == 2:
                                nonInseriti.append(dizio + ' ' + '11:50')
                            elif contaore == 3:
                                nonInseriti.append(dizio + ' ' + '13:40')
                            elif contaore == 4:
                                nonInseriti.append(dizio + ' ' + '15:00')
                            elif contaore == 5:
                                nonInseriti.append(dizio + ' ' + '16:40')
                            elif contaore == 6:
                                nonInseriti.append(dizio + ' ' + '18:00')
                            elif contaore == 7:
                                nonInseriti.append(dizio + ' ' + '19:20')
                            elif contaore == 8:
                                nonInseriti.append(dizio + ' ' + '20:40')

                    else:
                        if diz != u'':
                            if self.griglia.GetCellValue(1,contaore) == '' and inserito == False:
                                self.griglia.SetCellValue(1,contaore,unicode(dizio)+u' '+u'('+unicode(diz)+u')')
                                #inserito= True
                            elif self.griglia.GetCellValue(2, contaore) == '' and inserito == False:
                                self.griglia.SetCellValue(2, contaore,unicode(dizio) + u' ' + u'(' + unicode(diz) + u')')
                                #inserito = True
                            elif self.griglia.GetCellValue(3, contaore) == '' and inserito == False:
                                self.griglia.SetCellValue(3, contaore,
                                                          unicode(dizio) + u' ' + u'(' + unicode(diz) + u')')

                                #inserito = True
                            elif self.griglia.GetCellValue(4, contaore) == '' and inserito == False:
                                self.griglia.SetCellValue(4, contaore,
                                                          unicode(dizio) + u' ' + u'(' + unicode(diz) + u')')
                                #inserito = True
                            elif self.griglia.GetCellValue(6,contaore) == '' and inserito == False:
                                self.griglia.SetCellValue(6,contaore,unicode(dizio)+u' '+u'('+unicode(diz)+u')')
                                #inserito= True
                            elif self.griglia.GetCellValue(7, contaore) == '' and inserito == False:
                                self.griglia.SetCellValue(7, contaore,unicode(dizio) + u' ' + u'(' + unicode(diz) + u')')
                                #inserito = True
                            elif self.griglia.GetCellValue(8, contaore) == '' and inserito == False:
                                self.griglia.SetCellValue(8, contaore,
                                                          unicode(dizio) + u' ' + u'(' + unicode(diz) + u')')
                                #inserito = True
                            elif self.griglia.GetCellValue(9, contaore) == '' and inserito == False:
                                self.griglia.SetCellValue(9, contaore,
                                                          unicode(dizio) + u' ' + u'(' + unicode(diz) + u')')
                                #inserito = True
                            elif self.griglia.GetCellValue(11,contaore) == '' and inserito == False:
                                self.griglia.SetCellValue(11,contaore,unicode(dizio)+u' '+u'('+unicode(diz)+u')')
                                #inserito= True
                            elif self.griglia.GetCellValue(12, contaore) == '' and inserito == False:
                                self.griglia.SetCellValue(12, contaore, unicode(dizio) + u' ' + u'(' + unicode(diz) + u')')
                                #inserito = True
                            elif self.griglia.GetCellValue(13, contaore) == '' and inserito == False:
                                self.griglia.SetCellValue(13, contaore,
                                                           unicode(dizio) + u' ' + u'(' + unicode(diz) + u')')
                                #inserito = True
                            elif self.griglia.GetCellValue(14, contaore) == '' and inserito == False:
                                self.griglia.SetCellValue(14, contaore,
                                                         unicode(dizio) + u' ' + u'(' + unicode(diz) + u')')
                                #inserito = True
                            elif self.griglia.GetCellValue(16,contaore) == '' and inserito == False:
                                self.griglia.SetCellValue(16,contaore,unicode(dizio)+u' '+u'('+unicode(diz)+u')')
                                #inserito= True
                            elif self.griglia.GetCellValue(17, contaore) == '' and inserito == False:
                                self.griglia.SetCellValue(17, contaore, unicode(dizio) + u' ' + u'(' + unicode(diz) + u')')
                                #inserito = True
                            elif self.griglia.GetCellValue(18, contaore) == '' and inserito == False:
                                self.griglia.SetCellValue(18, contaore,
                                                         unicode(dizio) + u' ' + u'(' + unicode(diz) + u')')
                                #inserito = True
                            elif self.griglia.GetCellValue(19, contaore) == '' and inserito == False:
                                self.griglia.SetCellValue(19, contaore,
                                                          unicode(dizio) + u' ' + u'(' + unicode(diz) + u')')
                            elif self.griglia.GetCellValue(16, contaore) != '' and inserito == False:
                                if contaore == 0:
                                    nonInseriti.append(dizio + ' ' + '9:10')
                                elif contaore == 1:
                                    nonInseriti.append(dizio + ' ' + '10:30')
                                elif contaore == 2:
                                    nonInseriti.append(dizio + ' ' + '11:50')
                                elif contaore == 3:
                                    nonInseriti.append(dizio + ' ' + '13:40')
                                elif contaore == 4:
                                    nonInseriti.append(dizio + ' ' + '15:00')
                                elif contaore == 5:
                                    nonInseriti.append(dizio + ' ' + '16:40')
                                elif contaore == 6:
                                    nonInseriti.append(dizio + ' ' + '18:00')
                                elif contaore == 7:
                                    nonInseriti.append(dizio + ' ' + '19:20')
                                elif contaore == 8:
                                    nonInseriti.append(dizio + ' ' + '20:40')
                            #inserito = True

                    contaore = contaore + 1

        stringa = u'この生徒記入出来ません'
        for i in nonInseriti:
            stringa = unicode(stringa) + u' ' +unicode(i)
            print stringa

        d = wx.MessageDialog(None, stringa, '', wx.OK | wx.ICON_QUESTION)
        d.ShowModal()
        d.Destroy()

            # stanza1 = []
            # stanza2 = []
            # stanza3 = []
            # stanza4 = []
            # self.stanza1.Enabled = False
            # self.stanza2.Enabled = False
            # self.stanza3.Enabled = False
            # self.stanza4.Enabled = False
            # postiLiberi1 = 0
            # postiLiberi2 = 0
            # postiLiberi3 = 0
            # postiLiberi4 = 0
            # self.ore.DeselectAll()
            # global colonna
            # global riga
            # global colonnaSelezionata
            # global rigaSelezionata
            # for i in range(0, 9, 1):
            #     if self.materieVere.StringSelection == self.oreMaterie.GetCellValue(i, colonna):
            #         self.ore.Select(i)
            # if self.override.Value == True:
            #     self.listaMaterie.Clear()
            #     self.listaMaterie.Append(self.materieVere.StringSelection)
            # colonnaSelezionata = self.ore.Selections[0]
            # # [1, 2, 3, 4, 6, 7, 8, 9, 11, 12, 13, 14, 16, 17, 18, 19]
            # for i in range(0, self.griglia.NumberRows):
            #     if self.studentiDelGiorno.StringSelection in self.griglia.GetCellValue(i, colonnaSelezionata):
            #         self.stanza1.Enabled = False
            #         self.stanza2.Enabled = False
            #         self.stanza3.Enabled = False
            #         self.stanza4.Enabled = False
            #         stanza1 = []
            #         stanza2 = []
            #         stanza3 = []
            #         stanza4 = []
            #         print 'sonouscito'
            #         break
            #     if i >= 1 and i <= 4:
            #         stanza1.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            #     if i >= 6 and i <= 9:
            #         stanza2.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            #     if i >= 11 and i <= 14:
            #         stanza3.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            #     if i >= 16 and i <= 19:
            #         stanza4.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            # for i in stanza1:
            #     if i == unicode(''):
            #         postiLiberi1 = postiLiberi1 + 1
            # for i in stanza2:
            #     if i == u'':
            #         postiLiberi2 = postiLiberi2 + 1
            # for i in stanza3:
            #     if i == u'':
            #         postiLiberi3 = postiLiberi3 + 1
            # for i in stanza4:
            #     if i == u'':
            #         postiLiberi4 = postiLiberi4 + 1
            # print postiLiberi1, postiLiberi2, postiLiberi3, postiLiberi4
            # if postiLiberi1 >= 1:
            #     self.stanza1.Enabled = True
            # else:
            #     self.stanza1.Enabled = False
            # if postiLiberi2 >= 1:
            #     self.stanza2.Enabled = True
            # else:
            #     self.stanza2.Enabled = False
            # if postiLiberi3 >= 1:
            #     self.stanza3.Enabled = True
            # else:
            #     self.stanza3.Enabled = False
            # if postiLiberi4 >= 1:
            #     self.stanza4.Enabled = True
            # else:
            #     self.stanza4.Enabled = False
            # for i in stanza1:
            #     if 'K ' in i:
            #         self.stanza1.Enabled = False
            # for i in stanza2:
            #     if 'K ' in i:
            #         self.stanza2.Enabled = False
            # for i in stanza3:
            #     if 'K ' in i:
            #         self.stanza3.Enabled = False
            # for i in stanza4:
            #     if 'K ' in i:
            #         self.stanza4.Enabled = False
    def roomChange( self, event ):
        global colonnaSelezionata
        global rigaSelezionata
        global copia1
        global copia2
        global coordinateCopia1
        global coordinateCopia2
        global copia1m
        global copia2m
        global coordinateCopia1m
        global coordinateCopia2m
        global copia1Colore
        global copia2Colore
        global copia1Kojin
        global copia2Kojin
        colonnaSelezionata = event.GetCol()
        rigaSelezionata = event.GetRow()
        if rigaSelezionata == 0 or rigaSelezionata == 5 or rigaSelezionata == 10 or rigaSelezionata == 15:
            if copia1m != ' ' and copia2m == ' ':
                #copia2Colore = self.griglia.GetCellBackgroundColour(rigaSelezionata, colonnaSelezionata)
                copia2m = self.griglia.GetCellValue(rigaSelezionata, colonnaSelezionata)
                coordinateCopia2m = [rigaSelezionata, colonnaSelezionata]
                self.griglia.SetCellValue(rigaSelezionata, colonnaSelezionata, copia1m)
                self.griglia.SetCellValue(coordinateCopia1m[0], coordinateCopia1m[1], copia2m)
                copia1m = ' '
                copia2m = ' '
                coordinateCopia1m = []
                coordinateCopia2m = []
            if copia1m == ' ' and copia2m == ' ':
                copia1m = self.griglia.GetCellValue(rigaSelezionata, colonnaSelezionata)
                coordinateCopia1m = [rigaSelezionata, colonnaSelezionata]

            # if copia1 != ' ' and copia2 != ' ':
            #     # if copia1Kojin==True and copia2Kojin ==True:
            #     #copia2Colore == self.griglia.GetCellBackgroundColour(coordinateCopia1[0], coordinateCopia1[1])
            #     #copia1Colore == self.griglia.GetCellBackgroundColour(rigaSelezionata, colonnaSelezionata)
            #     self.griglia.SetCellBackgroundColour(coordinateCopia1[0], coordinateCopia1[1], copia2Colore)
            #     self.griglia.SetCellBackgroundColour(rigaSelezionata, colonnaSelezionata, copia1Colore)



        elif rigaSelezionata != 0 or rigaSelezionata !=5 or rigaSelezionata != 10 or rigaSelezionata !=15:
            if copia1 !=' ' and copia2 == ' ':
                copia2Colore =  self.griglia.GetCellBackgroundColour(rigaSelezionata,colonnaSelezionata)
                copia2 = self.griglia.GetCellValue(rigaSelezionata,colonnaSelezionata)
                coordinateCopia2 = [rigaSelezionata, colonnaSelezionata]
                self.griglia.SetCellValue(rigaSelezionata,colonnaSelezionata,copia1)
                self.griglia.SetCellValue(coordinateCopia1[0],coordinateCopia1[1],copia2)
                if self.griglia.GetCellValue(rigaSelezionata + 1, colonnaSelezionata) == u'(K)':
                    copia2Kojin = True

            if copia1 == ' ' and copia2 == ' ':
                copia1 = self.griglia.GetCellValue(rigaSelezionata,colonnaSelezionata)
                coordinateCopia1 = [rigaSelezionata,colonnaSelezionata]
                copia1Colore = self.griglia.GetCellBackgroundColour(rigaSelezionata,colonnaSelezionata)
                #self.griglia.SetCellBackgroundColour(rigaSelezionata,colonnaSelezionata,wx.BLUE)
                if self.griglia.GetCellValue(rigaSelezionata+1,colonnaSelezionata) == u'(K)':
                    copia1Kojin = True


            if copia1 != ' ' and copia2 != ' ':
                #if copia1Kojin==True and copia2Kojin ==True:
                copia2Colore == self.griglia.GetCellBackgroundColour(coordinateCopia1[0],coordinateCopia1[1])
                copia1Colore == self.griglia.GetCellBackgroundColour(rigaSelezionata,colonnaSelezionata)
                self.griglia.SetCellBackgroundColour(coordinateCopia1[0],coordinateCopia1[1],copia2Colore)
                self.griglia.SetCellBackgroundColour(rigaSelezionata,colonnaSelezionata, copia1Colore)
                copia1 = ' '
                copia2 = ' '
                coordinateCopia1 = []
                coordinateCopia2 = []
                copia1Colore = ''
                copia2Colore = ''
    def primaOraCheck( self, event ):
        acceso = self.primaCheck.IsChecked()
        print acceso, 'acceso'
        if acceso==1:
            global giorno
            materie = funzioni.cercaMaterie(1,giorno,tabella,tabellaTempo)
            self.kokugol.LabelText = str(materie[0])
            self.eigol.LabelText =  str(materie[1])
            self.suugakul.LabelText =  str(materie[2])
            self.rikal.LabelText =  str(materie[3])
            self.shakail.LabelText =  str(materie[4])
            self.tokubetsul.LabelText =  str(materie[5])
        if acceso==0:
            self.kokugol.LabelText = ''
            self.eigol.LabelText =  ''
            self.suugakul.LabelText =  ''
            self.rikal.LabelText =  ''
            self.shakail.LabelText = ''
            self.tokubetsul.LabelText =  ''
    def secondaOraCheck( self, event ):
        acceso = self.secondaCheck.IsChecked()
        print acceso, 'acceso'
        if acceso == 1:
            global giorno
            materie = funzioni.cercaMaterie(2, giorno, tabella, tabellaTempo)
            self.kokugol.LabelText = str(materie[0])
            self.eigol.LabelText = str(materie[1])
            self.suugakul.LabelText = str(materie[2])
            self.rikal.LabelText = str(materie[3])
            self.shakail.LabelText = str(materie[4])
            self.tokubetsul.LabelText = str(materie[5])
        if acceso == 0:
            self.kokugol.LabelText = ''
            self.eigol.LabelText = ''
            self.suugakul.LabelText = ''
            self.rikal.LabelText = ''
            self.shakail.LabelText = ''
            self.tokubetsul.LabelText = ''
    def terzaOraCheck( self, event ):
        acceso = self.terzaCheck.IsChecked()
        print acceso, 'acceso'
        if acceso == 1:
            global giorno
            materie = funzioni.cercaMaterie(3, giorno, tabella, tabellaTempo)
            self.kokugol.LabelText = str(materie[0])
            self.eigol.LabelText = str(materie[1])
            self.suugakul.LabelText = str(materie[2])
            self.rikal.LabelText = str(materie[3])
            self.shakail.LabelText = str(materie[4])
            self.tokubetsul.LabelText = str(materie[5])
        if acceso == 0:
            self.kokugol.LabelText = ''
            self.eigol.LabelText = ''
            self.suugakul.LabelText = ''
            self.rikal.LabelText = ''
            self.shakail.LabelText = ''
            self.tokubetsul.LabelText = ''
    def quartaOraCheck( self, event ):
        acceso = self.quartaCheck.IsChecked()
        print acceso, 'acceso'
        if acceso == 1:
            global giorno
            materie = funzioni.cercaMaterie(4, giorno, tabella, tabellaTempo)
            self.kokugol.LabelText = str(materie[0])
            self.eigol.LabelText = str(materie[1])
            self.suugakul.LabelText = str(materie[2])
            self.rikal.LabelText = str(materie[3])
            self.shakail.LabelText = str(materie[4])
            self.tokubetsul.LabelText = str(materie[5])
        if acceso == 0:
            self.kokugol.LabelText = ''
            self.eigol.LabelText = ''
            self.suugakul.LabelText = ''
            self.rikal.LabelText = ''
            self.shakail.LabelText = ''
            self.tokubetsul.LabelText = ''
    def quintaOraCheck( self, event ):
        acceso = self.quintaCheck.IsChecked()
        print acceso, 'acceso'
        if acceso == 1:
            global giorno
            materie = funzioni.cercaMaterie(5, giorno, tabella, tabellaTempo)
            self.kokugol.LabelText = str(materie[0])
            self.eigol.LabelText = str(materie[1])
            self.suugakul.LabelText = str(materie[2])
            self.rikal.LabelText = str(materie[3])
            self.shakail.LabelText = str(materie[4])
            self.tokubetsul.LabelText = str(materie[5])
        if acceso == 0:
            self.kokugol.LabelText = ''
            self.eigol.LabelText = ''
            self.suugakul.LabelText = ''
            self.rikal.LabelText = ''
            self.shakail.LabelText = ''
            self.tokubetsul.LabelText = ''
    def sestaOraCheck( self, event ):
        acceso = self.sestaCheck.IsChecked()
        print acceso, 'acceso'
        if acceso == 1:
            global giorno
            materie = funzioni.cercaMaterie(6, giorno, tabella, tabellaTempo)
            self.kokugol.LabelText = str(materie[0])
            self.eigol.LabelText = str(materie[1])
            self.suugakul.LabelText = str(materie[2])
            self.rikal.LabelText = str(materie[3])
            self.shakail.LabelText = str(materie[4])
            self.tokubetsul.LabelText = str(materie[5])
        if acceso == 0:
            self.kokugol.LabelText = ''
            self.eigol.LabelText = ''
            self.suugakul.LabelText = ''
            self.rikal.LabelText = ''
            self.shakail.LabelText = ''
            self.tokubetsul.LabelText = ''
    def settimaOraCheck( self, event ):
        acceso = self.settimaCheck.IsChecked()
        print acceso, 'acceso'
        if acceso == 1:
            global giorno
            materie = funzioni.cercaMaterie(7, giorno, tabella, tabellaTempo)
            self.kokugol.LabelText = str(materie[0])
            self.eigol.LabelText = str(materie[1])
            self.suugakul.LabelText = str(materie[2])
            self.rikal.LabelText = str(materie[3])
            self.shakail.LabelText = str(materie[4])
            self.tokubetsul.LabelText = str(materie[5])
        if acceso == 0:
            self.kokugol.LabelText = ''
            self.eigol.LabelText = ''
            self.suugakul.LabelText = ''
            self.rikal.LabelText = ''
            self.shakail.LabelText = ''
            self.tokubetsul.LabelText = ''
    def ottavaOraCheck( self, event ):
        acceso = self.ottavaCheck.IsChecked()
        print acceso, 'acceso'
        if acceso == 1:
            global giorno
            materie = funzioni.cercaMaterie(8, giorno, tabella, tabellaTempo)
            self.kokugol.LabelText = str(materie[0])
            self.eigol.LabelText = str(materie[1])
            self.suugakul.LabelText = str(materie[2])
            self.rikal.LabelText = str(materie[3])
            self.shakail.LabelText = str(materie[4])
            self.tokubetsul.LabelText = str(materie[5])
        if acceso == 0:
            self.kokugol.LabelText = ''
            self.eigol.LabelText = ''
            self.suugakul.LabelText = ''
            self.rikal.LabelText = ''
            self.shakail.LabelText = ''
            self.tokubetsul.LabelText = ''
    def nonaOraCheck( self, event ):
        acceso = self.nonaCheck.IsChecked()
        print acceso, 'acceso'
        if acceso == 1:
            global giorno
            materie = funzioni.cercaMaterie(9, giorno, tabella, tabellaTempo)
            self.kokugol.LabelText = str(materie[0])
            self.eigol.LabelText = str(materie[1])
            self.suugakul.LabelText = str(materie[2])
            self.rikal.LabelText = str(materie[3])
            self.shakail.LabelText = str(materie[4])
            self.tokubetsul.LabelText = str(materie[5])
        if acceso == 0:
            self.kokugol.LabelText = ''
            self.eigol.LabelText = ''
            self.suugakul.LabelText = ''
            self.rikal.LabelText = ''
            self.shakail.LabelText = ''
            self.tokubetsul.LabelText = ''
    def manualCheckOut( self, event ):
        global rigaSelezionata
        global colonnaSelezionata
        data = funzioni.aggiungizeri(self.calendario.Date.Year,self.calendario.Date.Month+1,self.calendario.Date.Day)

        ora = '00:00:00'
        tempo = data+' '+ora
        nomeVero = funzioni.ripuliscinome(self.griglia.GetCellValue(rigaSelezionata, colonnaSelezionata))
        if nomeVero is not None:
            tabellaTempo.insert(dict(name=nomeVero, time=tempo, giorno=data, ora=ora))
    def lezioniAggiuntive( self, event ):
        global colonnaSelezionata
        global rigaSelezionata
        studentedaElaborare = self.griglia.GetCellValue(rigaSelezionata, colonnaSelezionata)

        aggiungiAsterisco = funzioni.aggiungiAsterisco(self.griglia.GetCellValue(rigaSelezionata, colonnaSelezionata))
        if aggiungiAsterisco == True:
            self.griglia.SetCellValue(rigaSelezionata, colonnaSelezionata, studentedaElaborare.strip('*'))
        if aggiungiAsterisco == False:
            self.griglia.SetCellValue(rigaSelezionata, colonnaSelezionata, '*'+studentedaElaborare)
    def cancellaCelle( self, event ):
        global colonnaSelezionata
        global rigaSelezionata
        colonnaSelezionata = event.GetCol()
        rigaSelezionata = event.GetRow()
        if rigaSelezionata >= 0 and rigaSelezionata <= 4:
            indiceriga = 0
        if rigaSelezionata >=5  and rigaSelezionata <= 9:
            indiceriga = 5
        if rigaSelezionata >= 10 and rigaSelezionata <= 14:
            indiceriga = 10
        if rigaSelezionata >= 15 and rigaSelezionata <= 19:
            indiceriga = 15

        dlg = wx.MessageDialog(None, u"データ削除しますか", '', wx.YES_NO | wx.ICON_QUESTION)
        result = dlg.ShowModal()

        if result == wx.ID_YES:
            for i in range(0, 5):
                self.griglia.SetCellValue(indiceriga + i, colonnaSelezionata, '')
    def aggiungiStudentiAllaTabella(self,riga,colonna):
        global colonnaSelezionata
        global rigaSelezionata
        global stanza

        if rigaSelezionata >= 0 and rigaSelezionata <= 4:
            stanza = 0

        if rigaSelezionata >= 5 and rigaSelezionata <= 9:
            stanza = 5

        if rigaSelezionata >= 10 and rigaSelezionata <= 14:
            stanza = 10

        if rigaSelezionata >= 15 and rigaSelezionata <= 19:
            stanza = 15

        listaDaMandare = []

        # creare una lista con solo gli elementi che servono
        listaRigheStudenti = [1, 2, 3, 4, 6, 7, 8, 9, 11, 12, 13, 14, 16, 17, 18, 19]
        if stanza == 0:
            for i in range(6, 10, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            for i in range(11, 15, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            for i in range(16, 20, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        if stanza == 5:
            for i in range(1, 5, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            for i in range(11, 14, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            for i in range(16, 20, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        if stanza == 10:
            for i in range(6, 9, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            for i in range(1, 5, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            for i in range(16, 20, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        if stanza == 15:
            for i in range(6, 9, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            for i in range(11, 14, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            for i in range(1, 5, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        print listaDaMandare, 'lista da mandare'
        if rigaSelezionata >= 0 + stanza and rigaSelezionata <= 4 + stanza:
            # print 'stanza 1'
            # print self.listaStudenti.GetSelections(), type(self.listaStudenti.GetSelections())
            elementoget = self.listaStudenti.GetSelections()
            # print elementoget, 'elementoget'
            valoriStudenti = []
            print 'valore stanza', stanza
            if len(self.listaStudenti.GetSelections()) == 1:
                valoriStudenti.append(self.listaStudenti.Items[elementoget[0]])
                controlloDuplicati = funzioni.controlloDuplicatiStudenti(colonnaSelezionata, rigaSelezionata,
                                                                         valoriStudenti, listaDaMandare)

                if controlloDuplicati == True :
                    self.griglia.SetCellValue(1 + stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[0]])
                    self.griglia.SetCellValue(2 + stanza, colonnaSelezionata, u'')
                    self.griglia.SetCellValue(3 + stanza, colonnaSelezionata, u'')
                    self.griglia.SetCellValue(4 + stanza, colonnaSelezionata, u'')
                if controlloDuplicati == False:
                    dlg = wx.MessageDialog(None, u"生徒入力済み", '', wx.OK | wx.ICON_QUESTION)
                    result = dlg.ShowModal()
            if len(self.listaStudenti.GetSelections()) == 4:
                valoriStudenti.append(self.listaStudenti.Items[elementoget[0]])
                valoriStudenti.append(self.listaStudenti.Items[elementoget[1]])
                valoriStudenti.append(self.listaStudenti.Items[elementoget[2]])
                valoriStudenti.append(self.listaStudenti.Items[elementoget[3]])
                controlloDuplicati = funzioni.controlloDuplicatiStudenti(colonnaSelezionata, rigaSelezionata,
                                                                         valoriStudenti, listaDaMandare)
                if controlloDuplicati == True:
                    self.griglia.SetCellValue(1 + stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[0]])
                    self.griglia.SetCellValue(2 + stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[1]])
                    self.griglia.SetCellValue(3 + stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[2]])
                    self.griglia.SetCellValue(4 + stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[3]])
                if controlloDuplicati == False:
                    dlg = wx.MessageDialog(None, u"生徒入力済み", '', wx.YES_NO | wx.ICON_QUESTION)
                    result = dlg.ShowModal()
            if len(self.listaStudenti.GetSelections()) == 2:
                valoriStudenti.append(self.listaStudenti.Items[elementoget[0]])
                valoriStudenti.append(self.listaStudenti.Items[elementoget[1]])
                controlloDuplicati = funzioni.controlloDuplicatiStudenti(colonnaSelezionata, rigaSelezionata,
                                                                         valoriStudenti, listaDaMandare)
                if controlloDuplicati == True:
                    self.griglia.SetCellValue(1 + stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[0]])
                    self.griglia.SetCellValue(2 + stanza, colonnaSelezionata, '')
                    self.griglia.SetCellValue(3 + stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[1]])
                    self.griglia.SetCellValue(4 + stanza, colonnaSelezionata, '')
                if controlloDuplicati == False:
                    dlg = wx.MessageDialog(None, u"生徒入力済み", '', wx.YES_NO | wx.ICON_QUESTION)
                    result = dlg.ShowModal()
            if len(self.listaStudenti.GetSelections()) == 3:
                valoriStudenti.append(self.listaStudenti.Items[elementoget[0]])
                valoriStudenti.append(self.listaStudenti.Items[elementoget[1]])
                valoriStudenti.append(self.listaStudenti.Items[elementoget[2]])
                controlloDuplicati = funzioni.controlloDuplicatiStudenti(colonnaSelezionata, rigaSelezionata,
                                                                         valoriStudenti, listaDaMandare)
                if controlloDuplicati == True:
                    self.griglia.SetCellValue(1 + stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[0]])
                    self.griglia.SetCellValue(2 + stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[1]])
                    self.griglia.SetCellValue(3 + stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[2]])
                    self.griglia.SetCellValue(4 + stanza, colonnaSelezionata, '')
                if controlloDuplicati == False:
                    dlg = wx.MessageDialog(None, u"生徒入力済み", '', wx.YES_NO | wx.ICON_QUESTION)
                    result = dlg.ShowModal()
    def aggiungiStudentiAllaTabellaPerStanze(self,riga,colonna,elemento1,elemento2,elemento3,elemento4,nelementi):
        global colonnaSelezionata
        global rigaSelezionata
        global stanza

        if rigaSelezionata >= 0 and rigaSelezionata <= 4:
            stanza = 0

        if rigaSelezionata >= 5 and rigaSelezionata <= 9:
            stanza = 5

        if rigaSelezionata >= 10 and rigaSelezionata <= 14:
            stanza = 10

        if rigaSelezionata >= 15 and rigaSelezionata <= 19:
            stanza = 15

        listaDaMandare = []

        # creare una lista con solo gli elementi che servono
        listaRigheStudenti = [1, 2, 3, 4, 6, 7, 8, 9, 11, 12, 13, 14, 16, 17, 18, 19]
        if stanza == 0:
            for i in range(6, 10, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            for i in range(11, 15, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            for i in range(16, 20, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        if stanza == 5:
            for i in range(1, 5, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            for i in range(11, 14, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            for i in range(16, 20, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        if stanza == 10:
            for i in range(6, 9, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            for i in range(1, 5, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            for i in range(16, 20, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        if stanza == 15:
            for i in range(6, 9, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            for i in range(11, 14, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
            for i in range(1, 5, 1):
                listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        print listaDaMandare, 'lista da mandare'
        if rigaSelezionata >= 0 + stanza and rigaSelezionata <= 4 + stanza:
            # print 'stanza 1'
            # print self.listaStudenti.GetSelections(), type(self.listaStudenti.GetSelections())
            elementoget = self.listaStudenti.GetSelections()
            # print elementoget, 'elementoget'
            valoriStudenti = []
            print 'valore stanza', stanza

            # print self.griglia.GetCellValue(rigaSelezionata, colonnaSelezionata)


            #elf.griglia.SetCellValue(rigaSelezionata, colonnaSelezionata, pulisciStudente)
            if nelementi == 1:
                valoriStudenti.append(elemento1)
                controlloDuplicati = funzioni.controlloDuplicatiStudenti(colonnaSelezionata, rigaSelezionata,
                                                                         valoriStudenti, listaDaMandare)
                if controlloDuplicati == True:
                    self.griglia.SetCellValue(1 + stanza, colonnaSelezionata, elemento1)
                    self.griglia.SetCellValue(2 + stanza, colonnaSelezionata, u'')
                    self.griglia.SetCellValue(3 + stanza, colonnaSelezionata, u'')
                    self.griglia.SetCellValue(4 + stanza, colonnaSelezionata, u'')

                if controlloDuplicati == False:
                    dlg = wx.MessageDialog(None, u"生徒入力済み", '', wx.OK | wx.ICON_QUESTION)
                    result = dlg.ShowModal()
            if nelementi  == 4:
                valoriStudenti.append(elemento1)
                valoriStudenti.append(elemento2)
                valoriStudenti.append(elemento3)
                valoriStudenti.append(elemento4)
                controlloDuplicati = funzioni.controlloDuplicatiStudenti(colonnaSelezionata, rigaSelezionata,
                                                                         valoriStudenti, listaDaMandare)
                if controlloDuplicati == True:
                    self.griglia.SetCellValue(1 + stanza, colonnaSelezionata,elemento1)
                    self.griglia.SetCellValue(2 + stanza, colonnaSelezionata, elemento2)
                    self.griglia.SetCellValue(3 + stanza, colonnaSelezionata, elemento3)
                    self.griglia.SetCellValue(4 + stanza, colonnaSelezionata, elemento4)
                if controlloDuplicati == False:
                    dlg = wx.MessageDialog(None, u"生徒入力済み", '', wx.YES_NO | wx.ICON_QUESTION)
                    result = dlg.ShowModal()
            if nelementi == 2:
                valoriStudenti.append(elemento1)
                valoriStudenti.append(elemento2)
                controlloDuplicati = funzioni.controlloDuplicatiStudenti(colonnaSelezionata, rigaSelezionata,
                                                                         valoriStudenti, listaDaMandare)
                if controlloDuplicati == True:
                    self.griglia.SetCellValue(1 + stanza, colonnaSelezionata, elemento1)
                    self.griglia.SetCellValue(2 + stanza, colonnaSelezionata, elemento2)
                    self.griglia.SetCellValue(3 + stanza, colonnaSelezionata, '')
                    self.griglia.SetCellValue(4 + stanza, colonnaSelezionata, '')
                if controlloDuplicati == False:
                    dlg = wx.MessageDialog(None, u"生徒入力済み", '', wx.YES_NO | wx.ICON_QUESTION)
                    result = dlg.ShowModal()
            if nelementi == 3:
                valoriStudenti.append(elemento1)
                valoriStudenti.append(elemento2)
                valoriStudenti.append(elemento3)
                controlloDuplicati = funzioni.controlloDuplicatiStudenti(colonnaSelezionata, rigaSelezionata,
                                                                         valoriStudenti, listaDaMandare)
                if controlloDuplicati == True:
                    self.griglia.SetCellValue(1 + stanza, colonnaSelezionata, elemento1)
                    self.griglia.SetCellValue(2 + stanza, colonnaSelezionata, elemento2)
                    self.griglia.SetCellValue(3 + stanza, colonnaSelezionata, elemento3)
                    self.griglia.SetCellValue(4 + stanza, colonnaSelezionata, '')
                if controlloDuplicati == False:
                    dlg = wx.MessageDialog(None, u"生徒入力済み", '', wx.YES_NO | wx.ICON_QUESTION)
                    result = dlg.ShowModal()
    def selezionaStudenti(self,event):
        global colonnaSelezionata
        global rigaSelezionata
        global stanza
        self.aggiungiStudentiAllaTabella(rigaSelezionata,colonnaSelezionata)
        # global colonnaSelezionata
        # global rigaSelezionata
        # global stanza
        #
        # if rigaSelezionata >= 0 and rigaSelezionata <= 4:
        #     stanza = 0
        #
        # if rigaSelezionata >= 5 and rigaSelezionata <= 9:
        #     stanza = 5
        #
        # if rigaSelezionata >= 10 and rigaSelezionata <= 14:
        #     stanza = 10
        #
        # if rigaSelezionata >= 15 and rigaSelezionata <= 19:
        #     stanza = 15
        #
        # listaDaMandare = []
        #
        # # creare una lista con solo gli elementi che servono
        # listaRigheStudenti = [1, 2, 3, 4, 6, 7, 8, 9, 11, 12, 13, 14,16,17,18,19]
        # if stanza == 0:
        #     for i in range(6, 10, 1):
        #         listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        #     for i in range(11, 15, 1):
        #         listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        #     for i in range(16,20,1):
        #         listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        # if stanza == 5:
        #     for i in range(1, 5, 1):
        #         listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        #     for i in range(11, 14, 1):
        #         listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        #     for i in range(16, 20, 1):
        #         listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        # if stanza == 10:
        #     for i in range(6, 9, 1):
        #         listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        #     for i in range(1, 5, 1):
        #         listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        #     for i in range(16, 20, 1):
        #         listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        # if stanza == 15:
        #     for i in range(6, 9, 1):
        #         listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        #     for i in range(11, 14, 1):
        #         listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        #     for i in range(1, 5, 1):
        #         listaDaMandare.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        # print listaDaMandare, 'lista da mandare'
        # if rigaSelezionata >= 0+stanza and rigaSelezionata <= 4+stanza:
        #     # print 'stanza 1'
        #     # print self.listaStudenti.GetSelections(), type(self.listaStudenti.GetSelections())
        #     elementoget = self.listaStudenti.GetSelections()
        #     # print elementoget, 'elementoget'
        #     valoriStudenti = []
        #     print 'valore stanza',stanza
        #     if len(self.listaStudenti.GetSelections()) == 1:
        #         valoriStudenti.append(self.listaStudenti.Items[elementoget[0]])
        #         controlloDuplicati= funzioni.controlloDuplicatiStudenti(colonnaSelezionata,rigaSelezionata,valoriStudenti,listaDaMandare)
        #         if controlloDuplicati == True:
        #             self.griglia.SetCellValue(1+stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[0]])
        #             self.griglia.SetCellValue(2+stanza, colonnaSelezionata, u'個人')
        #             self.griglia.SetCellValue(3+stanza, colonnaSelezionata, u'個人')
        #             self.griglia.SetCellValue(4+stanza, colonnaSelezionata, u'個人')
        #         if controlloDuplicati == False:
        #             dlg = wx.MessageDialog(None, u"生徒入力済み", '', wx.OK | wx.ICON_QUESTION)
        #             result = dlg.ShowModal()
        #     if len(self.listaStudenti.GetSelections()) == 4:
        #         valoriStudenti.append(self.listaStudenti.Items[elementoget[0]])
        #         valoriStudenti.append(self.listaStudenti.Items[elementoget[1]])
        #         valoriStudenti.append(self.listaStudenti.Items[elementoget[2]])
        #         valoriStudenti.append(self.listaStudenti.Items[elementoget[3]])
        #         controlloDuplicati = funzioni.controlloDuplicatiStudenti(colonnaSelezionata, rigaSelezionata,
        #                                                                 valoriStudenti, listaDaMandare)
        #         if controlloDuplicati == True:
        #             self.griglia.SetCellValue(1+stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[0]])
        #             self.griglia.SetCellValue(2+stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[1]])
        #             self.griglia.SetCellValue(3+stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[2]])
        #             self.griglia.SetCellValue(4+stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[3]])
        #         if controlloDuplicati == False:
        #             dlg = wx.MessageDialog(None, u"生徒入力済み", '', wx.YES_NO | wx.ICON_QUESTION)
        #             result = dlg.ShowModal()
        #     if len(self.listaStudenti.GetSelections()) == 2:
        #         valoriStudenti.append(self.listaStudenti.Items[elementoget[0]])
        #         valoriStudenti.append(self.listaStudenti.Items[elementoget[1]])
        #         controlloDuplicati = funzioni.controlloDuplicatiStudenti(colonnaSelezionata, rigaSelezionata,
        #                                                                  valoriStudenti, listaDaMandare)
        #         if controlloDuplicati == True:
        #             self.griglia.SetCellValue(1+stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[0]])
        #             self.griglia.SetCellValue(2+stanza, colonnaSelezionata, '')
        #             self.griglia.SetCellValue(3+stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[1]])
        #             self.griglia.SetCellValue(4+stanza, colonnaSelezionata, '')
        #         if controlloDuplicati == False:
        #             dlg = wx.MessageDialog(None, u"生徒入力済み", '', wx.YES_NO | wx.ICON_QUESTION)
        #             result = dlg.ShowModal()
        #     if len(self.listaStudenti.GetSelections()) == 3:
        #         valoriStudenti.append(self.listaStudenti.Items[elementoget[0]])
        #         valoriStudenti.append(self.listaStudenti.Items[elementoget[1]])
        #         valoriStudenti.append(self.listaStudenti.Items[elementoget[2]])
        #         controlloDuplicati = funzioni.controlloDuplicatiStudenti(colonnaSelezionata, rigaSelezionata,
        #                                                                  valoriStudenti, listaDaMandare)
        #         if controlloDuplicati == True:
        #             self.griglia.SetCellValue(1+stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[0]])
        #             self.griglia.SetCellValue(2+stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[1]])
        #             self.griglia.SetCellValue(3+stanza, colonnaSelezionata, self.listaStudenti.Items[elementoget[2]])
        #             self.griglia.SetCellValue(4+stanza, colonnaSelezionata, '')
        #         if controlloDuplicati == False:
        #             dlg = wx.MessageDialog(None, u"生徒入力済み", '', wx.YES_NO | wx.ICON_QUESTION)
        #             result = dlg.ShowModal()
    def kojinsettei( self, event ):
        global rigaSelezionata
        global colonnaSelezionata
        if self.griglia.GetCellValue(rigaSelezionata,colonnaSelezionata)!= u'':
            stringaDaRipulire = self.griglia.GetCellValue(rigaSelezionata,colonnaSelezionata)
            if 'K 'in self.griglia.GetCellValue(rigaSelezionata,colonnaSelezionata):

                stringaRipulita = stringaDaRipulire.strip('K ')
                self.griglia.SetCellValue(rigaSelezionata, colonnaSelezionata,stringaRipulita)
            else:
                self.griglia.SetCellValue(rigaSelezionata, colonnaSelezionata, 'K '+self.griglia.GetCellValue(rigaSelezionata,colonnaSelezionata))
    def mandaAStanza1( self, event ):
        global colonnaSelezionata
        global rigaSelezionata
        rigaSelezionata =1
        colonnaSelezionata = self.ore.Selections[0]
        #self.aggiungiStudentiAllaTabella(1,self.ore.Selections[0])
        materiaConParentesi = ' (' + self.materieVere.StringSelection + ')'

        listafittizia = []
        listaStudenti = []
        listaStudentiPulita = []
        for i in range (1,5):
            listafittizia.append(self.griglia.GetCellValue(i,colonnaSelezionata))
        print listafittizia, 'listafit'
        for i in listafittizia:
            if i == u'':

                listaStudenti.append(self.studentiDelGiorno.StringSelection+materiaConParentesi)
                break
            if 'K ' in i:

                break
            else:
                listaStudenti.append(i)

        print len(listaStudenti)
        calcololunghezzaloop = 4-len(listaStudenti)
        for i in range(0,calcololunghezzaloop):
            listaStudenti.append(u'')
        print listaStudenti, 'listastudenti'

        contaelementibuoni = 0
        for i in listaStudenti:
            if i != u'':
                contaelementibuoni = contaelementibuoni+1
        #for i in listaStudentiPulita:
        #self.listaStudenti.SetSelection(self.studentiDelGiorno.Selections[0])
        self.aggiungiStudentiAllaTabellaPerStanze(rigaSelezionata,colonnaSelezionata, listaStudenti[0],listaStudenti[1],listaStudenti[2],listaStudenti[3],contaelementibuoni)
    def mandaAStanza2(self, event):
        global colonnaSelezionata
        global rigaSelezionata
        rigaSelezionata= 6

        colonnaSelezionata = self.ore.Selections[0]
        # self.aggiungiStudentiAllaTabella(1,self.ore.Selections[0])
        materiaConParentesi = ' (' + self.materieVere.StringSelection + ')'

        listafittizia = []
        listaStudenti = []
        listaStudentiPulita = []
        for i in range(6, 10):
            listafittizia.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        print listafittizia, 'listafit'
        for i in listafittizia:
            if i == u'':
                listaStudenti.append(self.studentiDelGiorno.StringSelection + materiaConParentesi)
                break
            if 'K ' in i:

                break
            else:
                listaStudenti.append(i)

        print len(listaStudenti)
        calcololunghezzaloop = 4 - len(listaStudenti)
        for i in range(0, calcololunghezzaloop):
            listaStudenti.append(u'')
        print listaStudenti, 'listastudenti'

        contaelementibuoni = 0
        for i in listaStudenti:
            if i != u'':
                contaelementibuoni = contaelementibuoni + 1
        # for i in listaStudentiPulita:
        # self.listaStudenti.SetSelection(self.studentiDelGiorno.Selections[0])
        self.aggiungiStudentiAllaTabellaPerStanze(rigaSelezionata, colonnaSelezionata, listaStudenti[0], listaStudenti[1],
                                                  listaStudenti[2], listaStudenti[3], contaelementibuoni)
    def mandaAStanza3(self, event):
        global colonnaSelezionata
        global rigaSelezionata
        rigaSelezionata= 11

        colonnaSelezionata = self.ore.Selections[0]
        # self.aggiungiStudentiAllaTabella(1,self.ore.Selections[0])
        materiaConParentesi = ' (' + self.materieVere.StringSelection + ')'

        listafittizia = []
        listaStudenti = []
        listaStudentiPulita = []
        for i in range(11, 15):
            listafittizia.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        print listafittizia, 'listafit'
        for i in listafittizia:
            if i == u'':
                listaStudenti.append(self.studentiDelGiorno.StringSelection + materiaConParentesi)
                break
            if 'K ' in i:

                break
            else:
                listaStudenti.append(i)

        print len(listaStudenti)
        calcololunghezzaloop = 4 - len(listaStudenti)
        for i in range(0, calcololunghezzaloop):
            listaStudenti.append(u'')
        print listaStudenti, 'listastudenti'

        contaelementibuoni = 0
        for i in listaStudenti:
            if i != u'':
                contaelementibuoni = contaelementibuoni + 1
        # for i in listaStudentiPulita:
        # self.listaStudenti.SetSelection(self.studentiDelGiorno.Selections[0])
        self.aggiungiStudentiAllaTabellaPerStanze(rigaSelezionata, colonnaSelezionata, listaStudenti[0], listaStudenti[1],
                                                  listaStudenti[2], listaStudenti[3], contaelementibuoni)
    def mandaAStanza4(self, event):
        global colonnaSelezionata
        global rigaSelezionata
        rigaSelezionata= 16
        colonnaSelezionata = self.ore.Selections[0]
        # self.aggiungiStudentiAllaTabella(1,self.ore.Selections[0])
        materiaConParentesi = ' (' + self.materieVere.StringSelection + ')'

        listafittizia = []
        listaStudenti = []
        listaStudentiPulita = []
        for i in range(16, 20):
            listafittizia.append(self.griglia.GetCellValue(i, colonnaSelezionata))
        print listafittizia, 'listafit'
        for i in listafittizia:
            if i == u'':
                listaStudenti.append(self.studentiDelGiorno.StringSelection + materiaConParentesi)
                break
            if 'K ' in i:

                break
            else:
                listaStudenti.append(i)

        print len(listaStudenti)
        calcololunghezzaloop = 4 - len(listaStudenti)
        for i in range(0, calcololunghezzaloop):
            listaStudenti.append(u'')
        print listaStudenti, 'listastudenti'

        contaelementibuoni = 0
        for i in listaStudenti:
            if i != u'':
                contaelementibuoni = contaelementibuoni + 1
        # for i in listaStudentiPulita:
        # self.listaStudenti.SetSelection(self.studentiDelGiorno.Selections[0])
        self.aggiungiStudentiAllaTabellaPerStanze(rigaSelezionata, colonnaSelezionata, listaStudenti[0], listaStudenti[1],
                                                  listaStudenti[2], listaStudenti[3], contaelementibuoni)
    def aggiungiMateria( self, event ):
        global colonnaSelezionata
        global rigaSelezionata
        #print self.griglia.GetCellValue(rigaSelezionata, colonnaSelezionata)
        materiaConParentesi = ' ('+self.listaMaterie.StringSelection+')'

        pulisciStudente = funzioni.puliscinome(self.griglia.GetCellValue(rigaSelezionata, colonnaSelezionata),materiaConParentesi)

        self.griglia.SetCellValue(rigaSelezionata, colonnaSelezionata, pulisciStudente)
        self.listaMaterie.DeselectAll()
    def scremaMateria(self, event):
        self.listaStudenti.Clear()

        popolaStudenti = tabella.find(student='1')
        global colonnaSelezionata
        # colonnaSelezionata = self.griglia.wxGridSelectCells
        listaAggiornataStudenti = funzioni.elaboraOraStudenti(tabella, colonnaSelezionata,
                                                              self.listaMaterie.StringSelection)

        for i in listaAggiornataStudenti:
            self.listaStudenti.Append(i)
    def scremaGente(self, event):
        popolaInsegnanti = tabellaIns.find(teacher='1')
        popolaStudenti = tabella.find(student='1')


        ottieniColonna = event.GetCol()
        global colonnaSelezionata
        colonnaSelezionata = event.GetCol()
        global rigaSelezionata
        rigaSelezionata = event.GetRow()
        print rigaSelezionata, 'riga selezionata', colonnaSelezionata, ' Colonna selezionata'
        ottieniRiga = event.GetRow()
        self.contenutoCella.LabelText=self.griglia.GetCellValue(rigaSelezionata,colonnaSelezionata)
        dataComposta = funzioni.aggiungizeri(self.calendario.Date.Year, self.calendario.Date.Month + 1,
                                             self.calendario.Date.Day)
        listaAggiornataInsegnanti = funzioni.elaboraOra(ottieniColonna, popolaInsegnanti, tabella,tabellaIns, tabellaTempoIns,dataComposta,tabellaDateIns, str(self.calendario.Date))
        #listaAggiornataStudenti=funzioni.elaboraOraStudenti(ottieniColonna,popolaStudenti,tabella,tabellaTempo,str(self.calendario.Date))
        self.listaInsegnanti.Clear()
        #self.listaStudenti.Clear()

        for i in listaAggiornataInsegnanti:
            self.listaInsegnanti.Append(i)
        #for i in listaAggiornataStudenti:
        #	self.listaStudenti.Append(i)
    def caricaCSV(self):
        self.listaStudenti.Clear()
        dataComposta = funzioni.aggiungizeri(self.calendario.Date.Year, self.calendario.Date.Month + 1,
                                             self.calendario.Date.Day)
        studentiElaborati = funzioni.elaboraOraStudenti(colonnaSelezionata, tabella, tabellaTempo, dataComposta,
                                                        self.listaMaterie.StringSelection)

        for i in studentiElaborati:
            controlloDuplicati = funzioni.controlloNomiDuplicati(i, self.listaStudenti.Items)
            if controlloDuplicati == True:
                self.listaStudenti.Append(i)
        print colonnaSelezionata, rigaSelezionata
        self.studentiDelGiorno.Clear()
        calendario = calendar
        global giorno
        giornoDelMeseCorrente = str(self.calendario.Date)
        dataDatetime = datetime.strptime(giornoDelMeseCorrente, '%m/%d/%y %H:%M:%S')
        lungezzaMese = calendario.monthrange(dataDatetime.year, dataDatetime.month)

        dataComp = str(self.calendario.Date.Year) + '/' + str(
            self.calendario.Date.Month + 1) + '/' + str(self.calendario.Date.Day)
        dataComposta = funzioni.aggiungizeri(self.calendario.Date.Year, self.calendario.Date.Month + 1,
                                             self.calendario.Date.Day)
        studentiPerData = tabellaTempo.find(uscite=dataComposta)
        # self.kokugoCheck.SetValue(0)
        # self.eigoCheck.SetValue(0)
        # self.suugakuCheck.SetValue(0)
        # self.rikaCheck.SetValue(0)
        # self.shakaiCheck.SetValue(0)
        # self.tokubetsuCheck.SetValue(0)
        self.dataText.LabelText = dataComposta

        giorno = dataComposta
        primaConta=0
        secondaConta = 0
        terzaConta = 0
        quartaConta = 0
        quintaConta = 0
        sestaConta = 0
        settimaConta = 0
        ottavaConta = 0
        nonaConta = 0
        kokugoConta = 0
        eigoConta = 0
        suugakuConta = 0
        rikaConta = 0
        shakaiConta = 0
        tokubetsuConta = 0

        cercaorariStudente = tabella.find(student='1')
        for i in cercaorariStudente:
            cercaStudentiDelGiorno = tabellaTempo.find_one(name=i.name, uscite =dataComposta)
            #print cercaStudentiDelGiorno.name
            if cercaStudentiDelGiorno is not None:
                self.studentiDelGiorno.Append(cercaStudentiDelGiorno.name)
        for i in studentiPerData:
            prima = tabella.count(name=i.name, primaOra=1)
            if prima == 1:
                primaConta = primaConta+prima
            seconda = tabella.count(name=i.name, secondaOra=1)
            if seconda == 1:
                secondaConta=secondaConta + seconda
            terza = tabella.count(name=i.name, terzaOra=1)
            if terza == 1:
                terzaConta = terzaConta + terza
            quarta = tabella.count(name=i.name, quartaOra=1)
            if quarta == 1:
                quartaConta = quartaConta + quarta
            quinta = tabella.count(name=i.name, quintaOra=1)
            if quinta == 1:
                quintaConta = quintaConta + quinta
            sesta = tabella.count(name=i.name, sestaOra=1)
            if sesta == 1:
                sestaConta = sestaConta + sesta
            settima = tabella.count(name=i.name, settimaOra=1)
            if settima == 1:
                settimaConta = settimaConta + settima
            ottava = tabella.count(name=i.name, ottavaOra=1)
            if ottava == 1:
                ottavaConta = ottavaConta + ottava
            nona = tabella.count(name=i.name, nonaOra=1)
            if nona == 1:
                nonaConta = nonaConta + nona
        for i in studentiPerData:
            kokugo = tabella.count(name=i.name, kokugo=1)
            if kokugo == 1:
                kokugoConta = kokugoConta+kokugo
            eigo = tabella.count(name=i.name, eigo=1)
            if eigo == 1:
                eigoConta = eigoConta + eigo
            suugaku = tabella.count(name=i.name, suugaku=1)
            if suugaku == 1:
                suugakuConta = suugakuConta + suugaku
            rika = tabella.count(name=i.name, rika=1)
            if rika == 1:
                rikaConta = rikaConta + rika
            shakai = tabella.count(name=i.name, shakai=1)
            if shakai == 1:
                shakaiConta = shakaiConta + shakai
            tokubetsu = tabella.count(name=i.name, tokubetsu=1)
            if tokubetsu == 1:
                tokubetsuConta = tokubetsuConta + tokubetsu
        # self.prima.LabelText = str(primaConta)
        # self.seconda.LabelText = str(secondaConta)
        # self.terza.LabelText = str(terzaConta)
        # self.quarta.LabelText = str(quartaConta)
        # self.quinta.LabelText = str(quintaConta)
        # self.sesta.LabelText = str(sestaConta)
        # self.settima.LabelText = str(settimaConta)
        # self.ottava.LabelText = str(ottavaConta)
        # self.nona.LabelText = str(nonaConta)



        nomeFile = str(self.calendario.Date)

        nomeFile = nomeFile.replace('/', '-')
        nomeFile = nomeFile.replace(' 00:00:00', '')
        anno = '20' + nomeFile[-2:]
        global percorso
        percorso = './csv/' + anno + '/' + nomeFile[:2] + '/' + nomeFile + '.csv'

        if not os.path.exists(os.path.dirname(percorso)):
            try:
                os.makedirs(os.path.dirname(percorso))
            except OSError as exc:  # Guard against race condition
                pass
        print percorso
        controllaPercorso = os.path.exists(percorso)

        if controllaPercorso == True:
            with open(percorso, 'rb') as f:
                reader = csv.DictReader(f)
                contarighe = 0
                converti = csvkit.unicsv.UnicodeCSVDictReader(f=f, encoding='utf-8')

                for i in converti:

                    self.griglia.SetCellValue(contarighe, 0, i['9:10 - 10:20'])
                    self.griglia.SetCellValue(contarighe, 1, i['10:30 - 11:40'])
                    self.griglia.SetCellValue(contarighe, 2, i['11:50 - 13:00'])
                    self.griglia.SetCellValue(contarighe, 3, i['13:40 - 14:50'])
                    self.griglia.SetCellValue(contarighe, 4, i['15:00 - 16:10'])
                    self.griglia.SetCellValue(contarighe, 5, i['16:40 - 17:50'])
                    self.griglia.SetCellValue(contarighe, 6, i['18:00 - 19:10'])
                    self.griglia.SetCellValue(contarighe, 7, i['19:20 - 20:30'])
                    self.griglia.SetCellValue(contarighe, 8, i['20:40 - 21:50'])
                    contarighe = contarighe + 1
        if controllaPercorso == False:
            contarighe = 0
            for i in range(0,20):
                self.griglia.SetCellValue(contarighe, 0, '')
                self.griglia.SetCellValue(contarighe, 1, '')
                self.griglia.SetCellValue(contarighe, 2, '')
                self.griglia.SetCellValue(contarighe, 3, '')
                self.griglia.SetCellValue(contarighe, 4, '')
                self.griglia.SetCellValue(contarighe, 5, '')
                self.griglia.SetCellValue(contarighe, 6, '')
                self.griglia.SetCellValue(contarighe, 7, '')
                self.griglia.SetCellValue(contarighe, 8, '')
                contarighe = contarighe + 1
        for r in range(1,5):
            for c in range (0,9):
               self.griglia.SetCellBackgroundColour(r, c, wx.WHITE)
        for r in range(6,10):
            for c in range (0,9):
               self.griglia.SetCellBackgroundColour(r, c, wx.WHITE)
        for r in range(11,15):
            for c in range (0,9):
               self.griglia.SetCellBackgroundColour(r, c, wx.WHITE)
        for r in range(16,20):
            for c in range (0,9):
               self.griglia.SetCellBackgroundColour(r, c, wx.WHITE)
        for r in range(1,5):
            for c in range (0,9):
                dataComposta = funzioni.aggiungizeri(self.calendario.Date.Year,
                                                     self.calendario.Date.Month + 1,
                                                     self.calendario.Date.Day)
                controlloCheckIn = funzioni.controlloCheckIn(self.griglia.GetCellValue(r, c),tabellaTempo,dataComposta)
                if controlloCheckIn == 'OUT':
                    self.griglia.SetCellBackgroundColour(r,c,wx.GREEN)
                if controlloCheckIn == "IN":
                    self.griglia.SetCellBackgroundColour(r,c,wx.RED)
                if controlloCheckIn == "NON":
                    self.griglia.SetCellBackgroundColour(r, c, wx.WHITE)
        for r in range(6,10):
            for c in range (0,9):
                dataComposta = funzioni.aggiungizeri(self.calendario.Date.Year,
                                                     self.calendario.Date.Month + 1,
                                                     self.calendario.Date.Day)
                controlloCheckIn = funzioni.controlloCheckIn(self.griglia.GetCellValue(r, c),tabellaTempo,dataComposta)
                if controlloCheckIn == 'OUT':
                    self.griglia.SetCellBackgroundColour(r,c,wx.GREEN)
                if controlloCheckIn == "IN":
                    self.griglia.SetCellBackgroundColour(r,c,wx.RED)
                if controlloCheckIn == "NON":
                    self.griglia.SetCellBackgroundColour(r, c, wx.WHITE)
        for r in range(11,15):
            for c in range (0,9):
                dataComposta = funzioni.aggiungizeri(self.calendario.Date.Year,
                                                     self.calendario.Date.Month + 1,
                                                     self.calendario.Date.Day)
                controlloCheckIn = funzioni.controlloCheckIn(self.griglia.GetCellValue(r, c),tabellaTempo,dataComposta)
                if controlloCheckIn == 'OUT':
                    self.griglia.SetCellBackgroundColour(r,c,wx.GREEN)
                if controlloCheckIn == "IN":
                    self.griglia.SetCellBackgroundColour(r,c,wx.RED)
                if controlloCheckIn == "NON":
                    self.griglia.SetCellBackgroundColour(r, c, wx.WHITE)
        for r in range(16,20):
            for c in range (0,9):
                dataComposta = funzioni.aggiungizeri(self.calendario.Date.Year,
                                                     self.calendario.Date.Month + 1,
                                                     self.calendario.Date.Day)
                controlloCheckIn = funzioni.controlloCheckIn(self.griglia.GetCellValue(r, c),tabellaTempo,dataComposta)
                if controlloCheckIn == 'OUT':
                    self.griglia.SetCellBackgroundColour(r,c,wx.GREEN)
                if controlloCheckIn == "IN":
                    self.griglia.SetCellBackgroundColour(r,c,wx.RED)
                if controlloCheckIn == "NON":
                    self.griglia.SetCellBackgroundColour(r, c, wx.WHITE)
    def preparativiInserimentoAutomatico(self,studente):
        valoriDaRestituire = []
        global colonna
        global riga
        self.materieVere.Clear()
        materieNecessarie = []
        colonna = 0
        riga = 0
        cercaStudente = studente
        studente = tabella.find_one(name=cercaStudente, student=1)
        # self.kokugoCheck.SetValue(studente.kokugo)
        # self.eigoCheck.SetValue(studente.eigo)
        # self.suugakuCheck.SetValue(studente.suugaku)
        # self.rikaCheck.SetValue(studente.rika)
        # self.shakaiCheck.SetValue(studente.shakai)
        # self.tokubetsuCheck.SetValue(studente.tokubetsu)
        self.individualCheck.SetValue(studente.individual)
        self.groupCheck.SetValue(studente.shared)

        popolastudenti = tabella.find_one(name=cercaStudente, student='1')

        percorsoStudenti = './StudentsData/' + popolastudenti.name + popolastudenti.telephone + '.txt'
        controllaPercorso = os.path.exists(percorsoStudenti)

        if controllaPercorso == True:
            with open(percorsoStudenti, 'rb') as f:
                reader = csv.DictReader(f)
                contarighe = 0
                converti = csvkit.unicsv.UnicodeCSVDictReader(f=f, encoding='utf-8')

                for i in converti:
                    self.oreMaterie.SetCellValue(contarighe, 0, i[u'月曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 1, i[u'火曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 2, i[u'水曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 3, i[u'木曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 4, i[u'金曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 5, i[u'土曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 6, i[u'日曜日'])

                    contarighe = contarighe + 1
        if controllaPercorso == False:
            self.oreMaterie.SelectAll()
            self.oreMaterie.ClearSelection()
            self.oreMaterie.ClearGrid()
            self.oreMaterie.Refresh()
            self.oreMaterie1.SelectAll()
            self.oreMaterie1.ClearSelection()
            self.oreMaterie1.ClearGrid()
            self.oreMaterie1.Refresh()

        if self.calendario.Date.WeekDay == 0:
            colonna = 6
        if self.calendario.Date.WeekDay == 1:
            colonna = 0
        if self.calendario.Date.WeekDay == 2:
            colonna = 1
        if self.calendario.Date.WeekDay == 3:
            colonna = 2
        if self.calendario.Date.WeekDay == 4:
            colonna = 3
        if self.calendario.Date.WeekDay == 5:
            colonna = 4
        if self.calendario.Date.WeekDay == 6:
            colonna = 5
        for i in range(0, 9, 1):

            valoriDaRestituire.append(self.oreMaterie.GetCellValue(i, colonna))
            print valoriDaRestituire, 'vloridarestituire'

        materieUniche = set(materieNecessarie)
        return valoriDaRestituire
        for i in materieUniche:
            self.materieVere.Append(i)


        percorsoStudenti1 = './StudentsData/' + popolastudenti.name + popolastudenti.telephone + 'tokubetsu.txt'
        controllaPercorso1 = os.path.exists(percorsoStudenti1)

        if controllaPercorso1 == True:
            with open(percorsoStudenti1, 'rb') as f1:
                reader1 = csv.DictReader(f1)
                contarighe1 = 0
                converti1 = csvkit.unicsv.UnicodeCSVDictReader(f=f1, encoding='utf-8')

                for i in converti1:
                    self.oreMaterie1.SetCellValue(contarighe1, 0, i[u'月曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 1, i[u'火曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 2, i[u'水曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 3, i[u'木曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 4, i[u'金曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 5, i[u'土曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 6, i[u'日曜日'])

                    contarighe1 = contarighe1 + 1
    def mostraStudentiDelGiorno( self, event ):
        global colonna
        global riga
        self.materieVere.Clear()
        materieNecessarie = []
        colonna=0
        riga = 0
        cercaStudente=self.studentiDelGiorno.StringSelection
        studente = tabella.find_one(name=cercaStudente, student=1)
        # self.kokugoCheck.SetValue(studente.kokugo)
        # self.eigoCheck.SetValue(studente.eigo)
        # self.suugakuCheck.SetValue(studente.suugaku)
        # self.rikaCheck.SetValue(studente.rika)
        # self.shakaiCheck.SetValue(studente.shakai)
        # self.tokubetsuCheck.SetValue(studente.tokubetsu)
        self.individualCheck.SetValue(studente.individual)
        self.groupCheck.SetValue(studente.shared)



        popolastudenti = tabella.find_one(name=self.studentiDelGiorno.StringSelection, student='1')

        percorsoStudenti = './StudentsData/' +popolastudenti.name + popolastudenti.telephone + '.txt'
        controllaPercorso = os.path.exists(percorsoStudenti)

        if controllaPercorso == True:
            with open(percorsoStudenti, 'rb') as f:
                reader = csv.DictReader(f)
                contarighe = 0
                converti = csvkit.unicsv.UnicodeCSVDictReader(f=f, encoding='utf-8')

                for i in converti:
                    self.oreMaterie.SetCellValue(contarighe, 0, i[u'月曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 1, i[u'火曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 2, i[u'水曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 3, i[u'木曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 4, i[u'金曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 5, i[u'土曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 6, i[u'日曜日'])

                    contarighe = contarighe + 1
        if controllaPercorso == False:
            self.oreMaterie.SelectAll()
            self.oreMaterie.ClearSelection()
            self.oreMaterie.ClearGrid()
            self.oreMaterie.Refresh()
            self.oreMaterie1.SelectAll()
            self.oreMaterie1.ClearSelection()
            self.oreMaterie1.ClearGrid()
            self.oreMaterie1.Refresh()

        if self.calendario.Date.WeekDay == 0:
            colonna = 6
        if self.calendario.Date.WeekDay == 1:
            colonna = 0
        if self.calendario.Date.WeekDay == 2:
            colonna = 1
        if self.calendario.Date.WeekDay == 3:
            colonna = 2
        if self.calendario.Date.WeekDay == 4:
            colonna = 3
        if self.calendario.Date.WeekDay == 5:
            colonna = 4
        if self.calendario.Date.WeekDay == 6:
            colonna = 5
        for i in range (0,9,1):
            if self.oreMaterie.GetCellValue(i,colonna)!= '':
                materieNecessarie.append(self.oreMaterie.GetCellValue(i,colonna))
        materieUniche = set(materieNecessarie)
        for i in materieUniche:
            self.materieVere.Append(i)

        percorsoStudenti1 = './StudentsData/' + popolastudenti.name + popolastudenti.telephone + 'tokubetsu.txt'
        controllaPercorso1 = os.path.exists(percorsoStudenti1)

        if controllaPercorso1 == True:
            with open(percorsoStudenti1, 'rb') as f1:
                reader1 = csv.DictReader(f1)
                contarighe1 = 0
                converti1 = csvkit.unicsv.UnicodeCSVDictReader(f=f1, encoding='utf-8')

                for i in converti1:
                    self.oreMaterie1.SetCellValue(contarighe1, 0, i[u'月曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 1, i[u'火曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 2, i[u'水曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 3, i[u'木曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 4, i[u'金曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 5, i[u'土曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 6, i[u'日曜日'])

                    contarighe1 = contarighe1 + 1
    def materieGiuste( self, event ):
        stanza1 = []
        stanza2 = []
        stanza3 = []
        stanza4 = []
        self.stanza1.Enabled = False
        self.stanza2.Enabled = False
        self.stanza3.Enabled = False
        self.stanza4.Enabled = False
        postiLiberi1 = 0
        postiLiberi2 = 0
        postiLiberi3 = 0
        postiLiberi4 = 0
        self.ore.DeselectAll()
        global colonna
        global riga
        global colonnaSelezionata
        global rigaSelezionata
        for i in range (0,9,1):
            if self.materieVere.StringSelection == self.oreMaterie.GetCellValue(i,colonna):
                self.ore.Select(i)
        if self.override.Value == True:
            self.listaMaterie.Clear()
            self.listaMaterie.Append(self.materieVere.StringSelection)
        colonnaSelezionata = self.ore.Selections[0]
        #[1, 2, 3, 4, 6, 7, 8, 9, 11, 12, 13, 14, 16, 17, 18, 19]
        for i in range(0,self.griglia.NumberRows):
            if self.studentiDelGiorno.StringSelection in self.griglia.GetCellValue(i,colonnaSelezionata):
                self.stanza1.Enabled = False
                self.stanza2.Enabled = False
                self.stanza3.Enabled = False
                self.stanza4.Enabled = False
                stanza1 = []
                stanza2 = []
                stanza3 = []
                stanza4 = []
                print 'sonouscito'
                break
            if i >=1 and i <=4:
                stanza1.append(self.griglia.GetCellValue(i,colonnaSelezionata))
            if i >=6 and i <=9:
                stanza2.append(self.griglia.GetCellValue(i,colonnaSelezionata))
            if i >=11 and i <=14:
                stanza3.append(self.griglia.GetCellValue(i,colonnaSelezionata))
            if i >=16 and i <=19:
                stanza4.append(self.griglia.GetCellValue(i,colonnaSelezionata))
        for i in stanza1:
            if i== unicode(''):
                postiLiberi1 = postiLiberi1+1
        for i in stanza2:
            if i== u'':
                postiLiberi2  = postiLiberi2+1
        for i in stanza3:
            if i== u'':
                postiLiberi3  = postiLiberi3+1
        for i in stanza4:
            if i== u'':
                postiLiberi4  = postiLiberi4+1
        print postiLiberi1,postiLiberi2,postiLiberi3,postiLiberi4
        if postiLiberi1 >=1 :
            self.stanza1.Enabled = True
        else:
            self.stanza1.Enabled = False
        if postiLiberi2 >=1:
            self.stanza2.Enabled = True
        else:
            self.stanza2.Enabled = False
        if postiLiberi3 >=1:
            self.stanza3.Enabled = True
        else:
            self.stanza3.Enabled = False
        if postiLiberi4 >=1:
            self.stanza4.Enabled = True
        else:
            self.stanza4.Enabled = False
        for i in stanza1:
            if 'K ' in i:
                self.stanza1.Enabled = False
        for i in stanza2:
            if 'K ' in i:
                self.stanza2.Enabled = False
        for i in stanza3:
            if 'K ' in i:
                self.stanza3.Enabled = False
        for i in stanza4:
            if 'K ' in i:
                self.stanza4.Enabled = False
    def materieSettateOverride( self, event ):
        if self.override.Value == True:
            self.listaMaterie.Clear()
            self.listaMaterie.Append(self.materieVere.StringSelection)
        if self.override.Value == False:
            self.listaMaterie.Clear()
            listaMaterie = [u'国語', u'英語', u'数学', u'理科', u'社会', u'特別']
            for i in listaMaterie:
                self.listaMaterie.Append(i)
    def inserisciInsegnante(self, event):

        global colonnaSelezionata
        global rigaSelezionata
        global stanza
        global materia
        print rigaSelezionata, type(rigaSelezionata)
        print self.griglia.GetCellValue(colonnaSelezionata, 0)
        controlloDuplicati = funzioni.controlloDuplicati(colonnaSelezionata, rigaSelezionata,
                                                         self.listaInsegnanti.StringSelection,
                                                         self.griglia.GetCellValue(0, colonnaSelezionata),
                                                         self.griglia.GetCellValue(5, colonnaSelezionata),
                                                         self.griglia.GetCellValue(10, colonnaSelezionata),
                                                         self.griglia.GetCellValue(15, colonnaSelezionata))
        if controlloDuplicati == True:
            #seleziona righe e stenze
            if rigaSelezionata >= 0 and rigaSelezionata <= 4:
                stanza = 0
                self.griglia.SetCellValue(0, colonnaSelezionata, self.listaInsegnanti.StringSelection)
                #self.griglia.SetCellValue(1, colonnaSelezionata, self.listaMaterie.StringSelection)
                #self.selezionaStudenti(self.listaMaterie.StringSelection,stanza)
            if rigaSelezionata >= 5 and rigaSelezionata <= 9:
                stanza = 5
                self.griglia.SetCellValue(5, colonnaSelezionata, self.listaInsegnanti.StringSelection)
                #self.griglia.SetCellValue(7, colonnaSelezionata, self.listaMaterie.StringSelection)
                #self.selezionaStudenti(self.listaMaterie.StringSelection,stanza)
            if rigaSelezionata >= 10 and rigaSelezionata <= 14:
                stanza = 10
                self.griglia.SetCellValue(10, colonnaSelezionata, self.listaInsegnanti.StringSelection)
                #self.griglia.SetCellValue(13, colonnaSelezionata, self.listaMaterie.StringSelection)
                #self.selezionaStudenti(self.listaMaterie.StringSelection,stanza)
            if rigaSelezionata >= 15 and rigaSelezionata <= 19:
                stanza = 15
                self.griglia.SetCellValue(15, colonnaSelezionata, self.listaInsegnanti.StringSelection)
                #self.griglia.SetCellValue(19, colonnaSelezionata, self.listaMaterie.StringSelection)
                #self.selezionaStudenti(self.listaMaterie.StringSelection,stanza)
        if controlloDuplicati == False:
            pass
    def selezionaStudentiEMaterie( self, event ):
        pass
    def selezionaMaterie(self, event):
        self.listaMaterie.Clear()
        materieSelezionateInsegnanti = funzioni.materieInsegnanti(self.listaInsegnanti.StringSelection, tabella)
        for i in materieSelezionateInsegnanti:
            self.listaMaterie.Append(i)
    def selezioneCalendario(self, event):
        with open(percorso, 'wb') as f:
            fieldnames = ['9:10 - 10:20', '10:30 - 11:40', '11:50 - 13:00', '13:40 - 14:50', '15:00 - 16:10',
                          '16:40 - 17:50', '18:00 - 19:10', '19:20 - 20:30', '20:40 - 21:50']
            writer = csv.DictWriter(f, fieldnames=fieldnames, dialect='excel')

            writer.writeheader()
            for i in range(0, 23, 1):
                #print i

                ciao = utf_8_encoder(self.griglia.GetCellValue(i, 0))
                #print ciao, 'ciao'
                writer.writerow(
                    {'9:10 - 10:20': utf_8_encoder(self.griglia.GetCellValue(i, 0)),
                     '10:30 - 11:40': utf_8_encoder(self.griglia.GetCellValue(i, 1))
                        , '11:50 - 13:00': utf_8_encoder(self.griglia.GetCellValue(i, 2)),
                     '13:40 - 14:50': utf_8_encoder(self.griglia.GetCellValue(i, 3))
                        , '15:00 - 16:10': utf_8_encoder(self.griglia.GetCellValue(i, 4)),
                     '16:40 - 17:50': utf_8_encoder(self.griglia.GetCellValue(i, 5))
                        , '18:00 - 19:10': utf_8_encoder(self.griglia.GetCellValue(i, 6)),
                     '19:20 - 20:30': utf_8_encoder(self.griglia.GetCellValue(i, 7))
                        , '20:40 - 21:50': utf_8_encoder(self.griglia.GetCellValue(i, 8))})
    def controlloGiornaliero( self, event ):
        Config = ConfigParser.ConfigParser()
        Config.read('./cfg.ini')
        colorestanza1 = Config.get('Colori stanze', 'colorestanza1')
        colorestanza2 = Config.get('Colori stanze', 'colorestanza2')
        colorestanza3 = Config.get('Colori stanze', 'colorestanza3')
        colorestanza4 = Config.get('Colori stanze', 'colorestanza4')
        listaMaterie = [u' (国語)', u' (英語)', u' (数学)', u' (理科)', u' (社会)', u' (特別)']
        for r in range(1, 5):
            for c in range(0, 9):
                if self.griglia.GetCellValue(r, c) != '':
                    for i in listaMaterie:
                        if i in self.griglia.GetCellValue(r, c) !=-1  :
                            self.griglia.SetCellBackgroundColour(r, c, wx.YELLOW)
        for r in range(6, 10):
            for c in range(0, 9):
                if self.griglia.GetCellValue(r, c) != '':
                    for i in listaMaterie:
                        if i in self.griglia.GetCellValue(r, c) !=-1  :
                            self.griglia.SetCellBackgroundColour(r, c, wx.YELLOW)
        for r in range(11, 15):
            for c in range(0, 9):
                if self.griglia.GetCellValue(r, c) != '':
                    for i in listaMaterie:
                        if i in self.griglia.GetCellValue(r, c) !=-1  :
                            self.griglia.SetCellBackgroundColour(r, c, wx.YELLOW)
        for r in range(16, 20):
            for c in range(0, 9):
                if self.griglia.GetCellValue(r, c) != '':
                    for i in listaMaterie:
                        if i in self.griglia.GetCellValue(r, c) !=-1  :
                            self.griglia.SetCellBackgroundColour(r, c, wx.YELLOW)
        self.griglia.Refresh()
        print 'aspetta'
    def salvaCSV(self, event):
        # self.primaCheck.SetValue(0)
        # self.secondaCheck.SetValue(0)
        # self.terzaCheck.SetValue(0)
        # self.quartaCheck.SetValue(0)
        # self.quintaCheck.SetValue(0)
        # self.sestaCheck.SetValue(0)
        # self.settimaCheck.SetValue(0)
        # self.ottavaCheck.SetValue(0)
        # self.nonaCheck.SetValue(0)
        # self.kokugol.LabelText=''
        # self.eigol.LabelText=''
        # self.suugakul.LabelText=''
        # self.rikal.LabelText=''
        # self.shakail.LabelText=''
        # self.tokubetsul.LabelText=''
        global percorso
        #self.caricaCSV()
        global datavecchia
        nomeFile = datavecchia

        nomeFile = nomeFile.replace('/', '-')
        nomeFile = nomeFile.replace(' 00:00:00', '')
        anno = '20' + nomeFile[-2:]
        percorso = './csv/' + anno + '/' + nomeFile[:2] + '/' + nomeFile + '.csv'


        with open(percorso, 'wb') as f:
            fieldnames = ['9:10 - 10:20', '10:30 - 11:40', '11:50 - 13:00', '13:40 - 14:50', '15:00 - 16:10',
                          '16:40 - 17:50', '18:00 - 19:10', '19:20 - 20:30', '20:40 - 21:50']
            writer = csv.DictWriter(f, fieldnames=fieldnames, dialect='excel')

            writer.writeheader()
            for i in range(0, 20, 1):
                #print i

                ciao = utf_8_encoder(self.griglia.GetCellValue(i, 0))

                writer.writerow(
                    {'9:10 - 10:20': utf_8_encoder(self.griglia.GetCellValue(i, 0)),
                     '10:30 - 11:40': utf_8_encoder(self.griglia.GetCellValue(i, 1))
                        , '11:50 - 13:00': utf_8_encoder(self.griglia.GetCellValue(i, 2)),
                     '13:40 - 14:50': utf_8_encoder(self.griglia.GetCellValue(i, 3))
                        , '15:00 - 16:10': utf_8_encoder(self.griglia.GetCellValue(i, 4)),
                     '16:40 - 17:50': utf_8_encoder(self.griglia.GetCellValue(i, 5))
                        , '18:00 - 19:10': utf_8_encoder(self.griglia.GetCellValue(i, 6)),
                     '19:20 - 20:30': utf_8_encoder(self.griglia.GetCellValue(i, 7))
                        , '20:40 - 21:50': utf_8_encoder(self.griglia.GetCellValue(i, 8))})

        #print datavecchia, 'datavecchiasalvacsv'
        datavecchia = str(self.calendario.Date)
        self.caricaCSV()
    def salvaDatiCSV(self, percorso):
        global datavecchia
        with open(percorso, 'wb') as f:
            fieldnames = ['9:10 - 10:20', '10:30 - 11:40', '11:50 - 13:00', '13:40 - 14:50', '15:00 - 16:10',
                          '16:40 - 17:50', '18:00 - 19:10', '19:20 - 20:30', '20:40 - 21:50']
            writer = csv.DictWriter(f, fieldnames=fieldnames, dialect='excel')

            writer.writeheader()
            for i in range(0, 23, 1):
                #print i

                ciao =  utf_8_encoder(self.griglia.GetCellValue(i, 0))
                #print ciao, 'ciao'
                writer.writerow(
                    {'9:10 - 10:20': utf_8_encoder(self.griglia.GetCellValue(i, 0)), '10:30 - 11:40':  utf_8_encoder(self.griglia.GetCellValue(i, 1))
                        , '11:50 - 13:00':  utf_8_encoder(self.griglia.GetCellValue(i, 2)),
                     '13:40 - 14:50':  utf_8_encoder(self.griglia.GetCellValue(i, 3))
                        , '15:00 - 16:10':  utf_8_encoder(self.griglia.GetCellValue(i, 4)),
                     '16:40 - 17:50':  utf_8_encoder(self.griglia.GetCellValue(i, 5))
                        , '18:00 - 19:10':  utf_8_encoder(self.griglia.GetCellValue(i, 6)),
                     '19:20 - 20:30':  utf_8_encoder(self.griglia.GetCellValue(i, 7))
                        , '20:40 - 21:50':  utf_8_encoder(self.griglia.GetCellValue(i, 8))})


        #print datavecchia, 'datavecchiasalvaDATIcsv'
        datavecchia = str(self.calendario.Date)
    # put a blank string in text when 'Clear' is clicked
    def clearFunc(self, event):
        self.text.SetValue(str(''))
    def FunzioneUserMenu(self, event):
        self.finestrautenti = finestraUtenti(None)
        self.finestrautenti.Show(True)
    def funzioneOpzioni( self, event ):
        self.finestraopzioni = finestraOpzioni(None)
        self.finestraopzioni.Show(True)

    def shiftinsegnanti(self, event):
        self.shiftInsegnanti = shiftinsegnanti(None)
        self.shiftInsegnanti.Show(True)

    def shiftstudenti(self, event):
        self.shiftStudenti = shiftstudenti(None)
        self.shiftStudenti.Show(True)

    def gestioneStudenti(self, event):
        self.finestrastudenti = finestraStudenti(None)
        self.finestrastudenti.Show(True)
    def mostraSalva( self, event ):
        self.salvadialog = saveDialog(None)
        self.salvadialog.Show(True)
    def mostraInfoStudente( self, event ):
        global mostraInfoStud
        global datiInfoStudente
        mostraInfoStud= True
        datiInfoStudente = self.studentiDelGiorno.StringSelection
        self.finestrastudenti = finestraStudenti(None)

        self.finestrastudenti.Show(True)
        datiInfoStudente = ''

class saveDialog(JukuPlanner.SaveDialog):
    def __init__(self, parent):
        # initialize parent class
        JukuPlanner.SaveDialog.__init__(self, parent)

class shiftstudenti(JukuPlanner.shiftGakusei):
    def __init__(self, parent):
        JukuPlanner.shiftGakusei.__init__(self, parent)
        popolastudenti = tabella.find(student='1')
        popolaInsegnanti = tabellaIns.find(teacher='1')
        for i in popolastudenti:
            self.listaSt.Append(unicode(i.name))
    def creaLoShift( self, event ):
        popolaInsegnanti = tabellaIns.find(teacher='1')
        #funzioni.generaShiftStudenti(self.listaSt.StringSelection,self.selezioneCartella.TextCtrlValue)
        funzioni.creashiftStudente(self.selezioneCartella.TextCtrlValue,self.listaSt.StringSelection,tabellaIns)
    def sendToOne( self, event ):
        popolastudenti = tabella.find_one(name=self.listaSt.StringSelection)
        funzioni.mandaShiftStudenti(popolastudenti.name,popolastudenti.email)
class finestraOpzioni(JukuPlanner.Opzioni):
    def __init__(self, parent):
        # initialize parent class
        JukuPlanner.Opzioni.__init__(self, parent)
        Config = ConfigParser.SafeConfigParser()
        Config.read('./cfg.ini')
        colorestanza1 = Config.get('Colori stanze', 'colorestanza1')
        colorestanza2 = Config.get('Colori stanze', 'colorestanza2')
        colorestanza3 = Config.get('Colori stanze', 'colorestanza3')
        colorestanza4 = Config.get('Colori stanze', 'colorestanza4')
        percorsoDB= Config.get('Paths','databasepath')
        percorsoCSV = Config.get('Paths', 'csvpath')
        percorsoStudenti = Config.get('Paths', 'studentpath')
        colore1 = funzioni.coonvertiStringaInColore(colorestanza1)
        colore2 = funzioni.coonvertiStringaInColore(colorestanza2)
        colore3 = funzioni.coonvertiStringaInColore(colorestanza3)
        colore4 = funzioni.coonvertiStringaInColore(colorestanza4)
        self.pickerstanza1.SetColour(wx.Colour(int(colore1[0]), int(colore1[1]), int(colore1[2]), int(colore1[3])))
        self.pickerstanza2.SetColour(wx.Colour(int(colore2[0]), int(colore2[1]), int(colore2[2]), int(colore2[3])))
        self.pickerstanza3.SetColour(wx.Colour(int(colore3[0]), int(colore3[1]), int(colore3[2]), int(colore3[3])))
        self.pickerstanza4.SetColour(wx.Colour(int(colore4[0]), int(colore4[1]), int(colore4[2]), int(colore4[3])))
        print percorsoDB
        self.percorsoDatabase.SetPath(percorsoDB)
        self.percorsoCSV.SetPath(percorsoCSV)
        self.percorsoStudenti.SetPath(percorsoStudenti)
        #self.mailLogin.Value = Config.get('Paths','databasePath')
    def settaColori( self, event ):
        Config = ConfigParser.ConfigParser()
        Config.add_section('Colori stanze')
        cfgfile = open("cfg.ini", 'w')
        Config.set('Colori stanze', 'ColoreStanza1', self.pickerstanza1.Colour)
        Config.set('Colori stanze', 'ColoreStanza2', self.pickerstanza2.Colour)
        Config.set('Colori stanze', 'ColoreStanza3', self.pickerstanza3.Colour)
        Config.set('Colori stanze', 'ColoreStanza4', self.pickerstanza4.Colour)
        Config.add_section('Paths')
        Config.set('Paths','databasePath',self.percorsoDatabase.TextCtrlValue)
        Config.set('Paths', 'csvpath', self.percorsoCSV.TextCtrlValue)
        Config.set('Paths', 'studentpath', self.percorsoStudenti.TextCtrlValue)
        Config.add_section('MailSetting')
        Config.set('MailSetting', 'login',self.mailLogin.Value)
        Config.set('MailSetting', 'password', self.mailPassword.Value)
        Config.set('MailSetting', 'server', self.mailServer.Value)

        Config.write(cfgfile)
        cfgfile.close()

    def caricaColori( self, event ):
        pass

class shiftinsegnanti(JukuPlanner.shiftSensei):
    def __init__(self, parent):
        JukuPlanner.shiftSensei.__init__(self,parent)
        popolaInsegnanti = tabellaIns.find(teacher='1')
        for i in popolaInsegnanti:
            self.listaIns.Append(unicode(i.name))
            print i.name
    def creaLoShift( self, event ):
        mailInsegnante = tabellaIns.find_one(teacher='1', name=self.listaIns.Selection)
        cartellaSelezionata = self.selezioneCartella.TextCtrlValue
        funzioni.creashift(cartellaSelezionata,self.listaIns.StringSelection)
    def sendToOne( self, event ):
        cartellaSelezionata = self.selezioneCartella.TextCtrlValue
        mailInsegnante = tabellaIns.find_one(teacher='1', name=self.listaIns.StringSelection)
        print mailInsegnante
        funzioni.mandaShift(mailInsegnante.name, mailInsegnante.email, self.linkDrive.Value)
class finestraStudenti(JukuPlanner.gakuseiFrame):
    def __init__(self, parent):
        # initialize parent class
        JukuPlanner.gakuseiFrame.__init__(self, parent)
        popolaStudenti = tabella.find(student='1')
        for i in popolaStudenti:
            self.listaStudenti.Append(unicode(i.name))
        self.oreMaterie.SetColMinimalWidth(0,30)
        self.grigliaTotali.SetCellValue(0, 0, '0')
        self.grigliaTotali.SetCellValue(0, 1, '0')
        self.grigliaTotali.SetCellValue(0, 2, '0')
        self.grigliaTotali.SetCellValue(0, 3, '0')
        self.grigliaTotali.SetCellValue(0, 4, '0')
        self.grigliaTotali.SetCellValue(0, 5, '0')
        self.grigliaTotali.SetCellValue(2, 0, '0')
        self.grigliaTotali.SetCellValue(2, 1, '0')
        self.grigliaTotali.SetCellValue(2, 2, '0')
        self.grigliaTotali.SetCellValue(2, 3, '0')
        self.grigliaTotali.SetCellValue(2, 4, '0')
        self.grigliaTotali.SetCellValue(2, 5, '0')
        self.nuovo.Enabled = False
        self.cancella.Enabled = False
        self.aggiorna.Enabled = False
        global contaPrivate
        global mostraInfoStud
        global datiInfoStudente
        if datiInfoStudente != '':
            self.listaStudenti.Clear()
            self.listaStudenti.Append(datiInfoStudente)
        contaPrivate = 0
    def cardDelete( self, event ):
        self.cardid.LabelText= ''
    def funzioneInvio(self, event):
        global contaPrivate
        contaPrivate = 0
        orario = {}
        orario1 = {}
        for creaorariofasullo in range(0, 9, 1):
            orario[creaorariofasullo] = False
        for creaorariofasullo in range(0, 9, 1):
            orario1[creaorariofasullo] = False
            #print orario[creaorariofasullo]
        cercaNome = tabella.find_one(name=self.casellaNome.Value)
        #print self.tabellaOre.Selections
        for i in self.tabellaOre.Selections:
            #print len(self.tabellaOre.Items)
            orario[i] = True
            #print orario[i]
        for i in self.tabellaOre1.Selections:
            #print len(self.tabellaOre.Items)
            orario1[i] = True
            #print orario[i]
        caselleDaCompletare = u''
        if self.casellaNome.Value == '':
            caselleDaCompletare = unicode(caselleDaCompletare)+ u'お名前入力してください. '
        if self.furigana.Value == '':
            caselleDaCompletare = unicode(caselleDaCompletare) + u'振り仮名入力してください. '
        if caselleDaCompletare != '':
            dlg = wx.MessageDialog(None, caselleDaCompletare, '', wx.ICON_QUESTION)
            result = dlg.ShowModal()
        if cercaNome is not None:
            self.errorCheck.LabelText = 'Name Already on database'

        if cercaNome is None and caselleDaCompletare == u'':
            self.errorCheck.LabelText = u'データ保存されました'
            tabellaGiorni.insert(dict(name=self.casellaNome.Value, lunedi=self.lunedi.Value,
                                     martedi=self.martedi.Value, mercoledi=self.mercoledi.Value,
                                     giovedi=self.giovedi.Value, venerdi=self.venerdi.Value,
                                     sabato=self.sabato.Value, domenica=self.domenica.Value,
                                      lunedi1=self.lunedi1.Value,
                                     martedi1=self.martedi1.Value, mercoledi1=self.mercoledi1.Value,
                                     giovedi1=self.giovedi1.Value, venerdi1=self.venerdi1.Value,
                                     sabato1=self.sabato1.Value, domenica1=self.domenica1.Value))
            tabella.insert(
                dict(name=self.casellaNome.Value, cardID=self.cardid.Label, telephone=self.casellaTelefono.Value,furigana = self.furigana.Value ,
                     email=self.casellaEmail.Value, parentMail=self.mailGenitori.Value, scuola=self.casellaScuola.Value,
                     maschio=self.danseiBox.Value, femmina=self.joseiBox.Value, student=1, sonota=self.sonota.Value,
                     teacher=0, kokugo=self.kokugo.Value, eigo=self.eigo.Value, suugaku=self.suugaku.Value,
                     rika=self.rika.Value, shakai=self.shakai.Value, tokubetsu=self.tokubetsu.Value,
                     primaOra=orario[0], secondaOra=orario[1], terzaOra=orario[2], quartaOra=orario[3],
                     quintaOra=orario[4], sestaOra=orario[5], settimaOra=orario[6], ottavaOra=orario[7],
                     nonaOra=orario[8],individual=self.individual.Value, shared=self.shared.Value,
                     kokugo1=self.kokugo1.Value, eigo1=self.eigo1.Value, suugaku1=self.suugaku1.Value,
                     rika1=self.rika1.Value, shakai1=self.shakai1.Value, tokubetsu1=self.tokubetsu1.Value,
                     primaOra1=orario1[0], secondaOra1=orario1[1], terzaOra1=orario1[2], quartaOra1=orario1[3],
                     quintaOra1=orario1[4], sestaOra1=orario1[5], settimaOra1=orario1[6], ottavaOra1=orario1[7],
                     nonaOra1=orario1[8],))
            tabellaCalcoli.insert(dict(name=self.casellaNome.Value,
                                       anno=self.calendarioStudenti.Date.Year ,
                                       mese=self.calendarioStudenti.Date.Month,
                                       normaleigo=self.grigliaTotali.GetCellValue(0,0),
                                       normalsuugaku=self.grigliaTotali.GetCellValue(0,1),
                                       normalkokugo=self.grigliaTotali.GetCellValue(0,2),
                                       normalrika=self.grigliaTotali.GetCellValue(0, 3),
                                       normalshakai=self.grigliaTotali.GetCellValue(0,4),
                                       normaltokubetsu=self.grigliaTotali.GetCellValue(0,5,),
                                       tsuikaeigo=self.grigliaTotali.GetCellValue(2,0),
                                       tsuikasuugaku=self.grigliaTotali.GetCellValue(2,1),
                                       tsuikakokugo=self.grigliaTotali.GetCellValue(2,2),
                                       tsuikarika=self.grigliaTotali.GetCellValue(2, 3),
                                       tsuikashakai=self.grigliaTotali.GetCellValue(2,4),
                                       tsuikatokubetsu=self.grigliaTotali.GetCellValue(2,5,)

                                       ))

            for i in self.usciteStudenti.Items:
                tabellaTempo.insert(dict(name=self.casellaNome.Value, uscite=i))

        self.listaStudenti.Clear()
        popolaStudenti = tabella.find(student='1')
        for i in popolaStudenti:
            self.listaStudenti.Append(unicode(i.name))
        self.oreMaterie.SetColMinimalWidth(0, 30)
        self.invio.Enabled = False

        #self.errorCheck.LabelText = u'データ保存されました'
    def materiePrivate( self, event ):
        global rigaMaterie
        global colonnaMaterie
        if self.oreMaterie.GetCellValue(rigaMaterie,colonnaMaterie)!= '':
            if self.oreMaterie.GetCellValue(rigaMaterie,colonnaMaterie)[0] == 'K':
                prestrippata = self.oreMaterie.GetCellValue(rigaMaterie,colonnaMaterie)
                strippata = prestrippata.strip('K')
                self.oreMaterie.SetCellValue(rigaMaterie, colonnaMaterie,
                                             strippata)
            else:
                self.oreMaterie.SetCellValue(rigaMaterie,colonnaMaterie, u'K'+ self.oreMaterie.GetCellValue(rigaMaterie,colonnaMaterie))
    def cancellaMaterie( self, event ):
        global rigaMaterie
        global colonnaMaterie
        rigaMaterie = event.GetRow()
        colonnaMaterie = event.GetCol()

        self.oreMaterie.SetCellValue(rigaMaterie, colonnaMaterie, '')
    def inviaShift( self, event ):
        random_data = os.urandom(128)
        nomerandom = hashlib.md5(random_data).hexdigest()[:16]
        shiftTemp = open('./shift/'+nomerandom+'.txt','w')
        datavecchia = str(self.calendarioStudenti.Date)
        nomeFile = datavecchia
        lezioniPrivate = 0
        nomeFile = nomeFile.replace('/', '-')
        nomeFile = nomeFile.replace(' 00:00:00', '')
        anno = '20' + nomeFile[-2:]
        percorso = './csv/' + anno + '/' + nomeFile[:2] + '/'

        files = os.listdir(percorso)

        files_txt = [i for i in files if i.endswith('.csv')]
        print files_txt
        for i in files_txt:
            with open(percorso + i, 'rb') as f:
                reader = csv.DictReader(f)
                contarighe = 0
                converti = csvkit.unicsv.UnicodeCSVDictReader(f=f, encoding='utf-8')
                print converti, converti

                for i in converti:
                    #if i['9:10 - 10:20'] == '':
                    #shiftTemp.write(i+ '9:10 - 10:20' + )

                    if i['10:30 - 11:40'] == self.listaStudenti.StringSelection:
                        lezioniPrivate = lezioniPrivate + 1
                    if i['11:50 - 13:00'] == self.listaStudenti.StringSelection:
                        lezioniPrivate = lezioniPrivate + 1
                    if i['13:40 - 14:50'] == self.listaStudenti.StringSelection:
                        lezioniPrivate = lezioniPrivate + 1
                    if i['15:00 - 16:10'] == self.listaStudenti.StringSelection:
                        lezioniPrivate = lezioniPrivate + 1
                    if i['18:00 - 19:10'] == self.listaStudenti.StringSelection:
                        lezioniPrivate = lezioniPrivate + 1
                    if i['19:20 - 20:30'] == self.listaStudenti.StringSelection:
                        lezioniPrivate = lezioniPrivate + 1
                    if i['20:40 - 21:50'] == self.listaStudenti.StringSelection:
                        lezioniPrivate = lezioniPrivate + 1

                    if i['9:10 - 10:20'] == str(self.listaStudenti.StringSelection):
                    #shiftTemp.write(i+ '9:10 - 10:20' + )
                        pass
                    if i['10:30 - 11:40'] == self.listaStudenti.StringSelection:
                        lezioniPrivate = lezioniPrivate + 1
                    if i['11:50 - 13:00'] == self.listaStudenti.StringSelection:
                        lezioniPrivate = lezioniPrivate + 1
                    if i['13:40 - 14:50'] == self.listaStudenti.StringSelection:
                        lezioniPrivate = lezioniPrivate + 1
                    if i['15:00 - 16:10'] == self.listaStudenti.StringSelection:
                        lezioniPrivate = lezioniPrivate + 1
                    if i['18:00 - 19:10'] == self.listaStudenti.StringSelection:
                        lezioniPrivate = lezioniPrivate + 1
                    if i['19:20 - 20:30'] == self.listaStudenti.StringSelection:
                        lezioniPrivate = lezioniPrivate + 1
                    if i['20:40 - 21:50'] == self.listaStudenti.StringSelection:
                        lezioniPrivate = lezioniPrivate + 1
                print  lezioniPrivate, 'lezioni private'
    def materieGiorno( self, event ):
        global rigaMaterie
        global colonnaMaterie

        print rigaMaterie,colonnaMaterie
        if len(self.materieGiorni.Items)==1:
            self.oreMaterie.SetCellValue(rigaMaterie, colonnaMaterie,self.materieGiorni.StringSelection)
            self.materieGiorni.DeselectAll()
        if len(self.materieGiorni.Items) > 1:
            self.oreMaterie.SetCellValue(rigaMaterie, colonnaMaterie, self.materieGiorni.StringSelection)
            self.materieGiorni.DeselectAll()
    def materieGiorno1( self, event ):
        global rigaMaterie1
        global colonnaMaterie1

        #print rigaMaterie,colonnaMaterie
        if len(self.materieGiorni1.Items)==1:
            self.oreMaterie1.SetCellValue(rigaMaterie1, colonnaMaterie1,self.materieGiorni1.StringSelection)
            self.materieGiorni1.DeselectAll()
        if len(self.materieGiorni1.Items) > 1:
            self.oreMaterie1.SetCellValue(rigaMaterie1, colonnaMaterie1, self.materieGiorni1.StringSelection)
            self.materieGiorni1.DeselectAll()
    def nuovoStudente( self, event ):

        self.grigliaLezioniSingole.SelectAll()
        self.grigliaLezioniSingole.ClearSelection()
        self.grigliaLezioniSingole.ClearGrid()
        self.grigliaLezioniSingole.Refresh()
        self.grigliaTotali.SelectAll()
        self.grigliaTotali.ClearSelection()
        self.grigliaTotali.ClearGrid()
        self.grigliaTotali.Refresh()
        for i in range(0,6,1):
            self.grigliaTotali.SetCellBackgroundColour(3,i,wx.WHITE)
        self.invio.Enabled=True
        self.aggiorna.Enabled=False
        self.cancella.Enabled=False
        self.casellaNome.Clear()
        self.casellaTelefono.Clear()
        self.furigana.Clear()
        self.danseiBox.Value = False
        self.joseiBox.Value = False
        self.sonota.Value = False
        self.casellaEmail.Clear()
        self.mailGenitori.Clear()
        self.casellaScuola.Clear()
        self.tabellaOre.DeselectAll()
        self.usciteStudenti.Clear()
        self.materieGiorni.Clear()
        self.kokugo.Value = False
        self.eigo.Value = False
        self.suugaku.Value = False
        self.rika.Value = False
        self.tokubetsu.Value = False
        self.shakai.Value = False
        self.oreMaterie.ClearGrid()
        self.lunedi.Value = False
        self.martedi.Value = False
        self.mercoledi.Value = False
        self.giovedi.Value = False
        self.venerdi.Value = False
        self.sabato.Value = False
        self.domenica.Value = False

        self.tabellaOre1.DeselectAll()
        self.materieGiorni1.Clear()
        self.kokugo1.Value = False
        self.eigo1.Value = False
        self.suugaku1.Value = False
        self.rika1.Value = False
        self.tokubetsu1.Value = False
        self.shakai1.Value = False
        self.oreMaterie1.ClearGrid()
        self.lunedi1.Value = False
        self.martedi1.Value = False
        self.mercoledi1.Value = False
        self.giovedi1.Value = False
        self.venerdi1.Value = False
        self.sabato1.Value = False
        self.domenica1.Value = False
        self.individual.Value = False
        self.shared.Value = False
        self.cardid.LabelText=''
        self.cardcancel.Enabled = False
        self.CardRegistration.Enabled = True
        for i in range(0, 9, 1):
            #self.griglia.SetCellBackgroundColour(0, i, wx.GREEN)
            self.oreMaterie1.SetCellBackgroundColour(0, i,wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(1, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(2, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(3, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(4, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(5, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(6, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(7, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(8, i, wx.WHITE)
        for i in range(0, 9, 1):
            #self.griglia.SetCellBackgroundColour(0, i, wx.GREEN)
            self.oreMaterie.SetCellBackgroundColour(0, i,wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(1, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(2, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(3, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(4, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(5, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(6, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(7, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(8, i, wx.WHITE)
        self.listaStudenti.DeselectAll()
    def femmina( self, event ):
        self.danseiBox.Value = False
        self.sonota.Value = False
    def maschio( self, event ):
        self.joseiBox.Value = False
        self.sonota.Value = False
    def mostraMeseCorrente( self, event ):
        listadate = []
        dataComp = str(self.calendarioStudenti.Date.Year) + '/' + str(self.calendarioStudenti.Date.Month + 1) + '/'
        dataComposta = funzioni.aggiungizeriSenzaGiorno(self.calendarioStudenti.Date.Year, self.calendarioStudenti.Date.Month + 1)
        dataunicode = unicode(dataComposta)
        contaitem = 0
        popolaDate = tabellaTempo.find(name=self.listaStudenti.StringSelection)

        if self.meseCorrente.Value == True:
            for i in self.usciteStudenti.Items:
                if dataunicode in i :
                    listadate.append(i)
            self.usciteStudenti.Clear()
            for i in listadate:
                self.usciteStudenti.Append(i)
        if self.meseCorrente.Value == False:
            self.usciteStudenti.Clear()
            for i in popolaDate:
                if len((str(i.uscite))) >= 5:
                    self.usciteStudenti.Append(str(i.uscite))
    def lgbt( self, event ):
        self.joseiBox.Value = False
        self.danseiBox.Value = False
    def selezionaCellaMateria( self, event ):
        global rigaMaterie
        global colonnaMaterie
        rigaMaterie = event.GetRow()
        colonnaMaterie = event.GetCol()
        self.oreMaterie1.SetCellValue(rigaMaterie1, colonnaMaterie1,self.materieGiorni1.StringSelection)
    def calcoliDiFineMese( self, event ):
        dlg = wx.MessageDialog(None, u"選択された生徒月末処理しますか", '', wx.YES_NO | wx.ICON_QUESTION)
        result = dlg.ShowModal()

        if result == wx.ID_YES:
            popolaCalcoli = tabellaCalcoli.find_one(name=self.listaStudenti.StringSelection,
                                                    anno=self.calendarioStudenti.Date.Year,
                                                    mese=self.calendarioStudenti.Date.Month)
            if popolaCalcoli is not None:
                datiCalcoli = (
                dict(id=popolaCalcoli.id, name=self.casellaNome.Value, mese=self.calendarioStudenti.Date.Month,
                     anno=self.calendarioStudenti.Date.Year, normaleigo=self.grigliaTotali.GetCellValue(0, 0),
                     normalsuugaku=self.grigliaTotali.GetCellValue(0, 1),
                     normalkokugo=self.grigliaTotali.GetCellValue(0, 2),
                     normalrika=self.grigliaTotali.GetCellValue(0, 3),
                     normalshakai=self.grigliaTotali.GetCellValue(0, 4),
                     normaltokubetsu=self.grigliaTotali.GetCellValue(0, 5),
                     tsuikaeigo=self.grigliaTotali.GetCellValue(2, 0),
                     tsuikasuugaku=self.grigliaTotali.GetCellValue(2, 1),
                     tsuikakokugo=self.grigliaTotali.GetCellValue(2, 2),
                     tsuikarika=self.grigliaTotali.GetCellValue(2, 3),
                     tsuikashakai=self.grigliaTotali.GetCellValue(2, 4),
                     tsuikatokubetsu=self.grigliaTotali.GetCellValue(2, 5),
                     balanceeigo=self.grigliaLezioniSingole.GetCellValue(33, 0),
                     balancesuugaku=self.grigliaLezioniSingole.GetCellValue(33, 1),
                     balancekokugo=self.grigliaLezioniSingole.GetCellValue(33, 2),
                     balancerika=self.grigliaLezioniSingole.GetCellValue(33, 3),
                     balanceshakai=self.grigliaLezioniSingole.GetCellValue(33, 4),
                     balancetokubetu=self.grigliaLezioniSingole.GetCellValue(33, 5)))

                tabellaCalcoli.update(datiCalcoli, ['id'])
            if popolaCalcoli is None:
                tabellaCalcoli.insert(dict(name=self.casellaNome.Value, anno=self.calendarioStudenti.Date.Year,
                                           mese=self.calendarioStudenti.Date.Month,
                                           normaleigo=self.grigliaTotali.GetCellValue(0, 0),
                                           normalsuugaku=self.grigliaTotali.GetCellValue(0, 1),
                                           normalkokugo=self.grigliaTotali.GetCellValue(0, 2),
                                           normalrika=self.grigliaTotali.GetCellValue(0, 3),
                                           normalshakai=self.grigliaTotali.GetCellValue(0, 4),
                                           normaltokubetsu=self.grigliaTotali.GetCellValue(0, 5),
                                           tsuikaeigo=self.grigliaTotali.GetCellValue(2, 0),
                                           tsuikasuugaku=self.grigliaTotali.GetCellValue(2, 1),
                                           tsuikakokugo=self.grigliaTotali.GetCellValue(2, 2),
                                           tsuikarika=self.grigliaTotali.GetCellValue(2, 3),
                                           tsuikashakai=self.grigliaTotali.GetCellValue(2, 4),
                                           tsuikatokubetsu=self.grigliaTotali.GetCellValue(2, 5),
                                           balanceeigo=self.grigliaLezioniSingole.GetCellValue(33, 0),
                                           balancesuugaku=self.grigliaLezioniSingole.GetCellValue(33, 1),
                                           balancekokugo=self.grigliaLezioniSingole.GetCellValue(33, 2),
                                           balancerika=self.grigliaLezioniSingole.GetCellValue(33, 3),
                                           balanceshakai=self.grigliaLezioniSingole.GetCellValue(33, 4),
                                           balancetokubetu=self.grigliaLezioniSingole.GetCellValue(33, 5)))
    def aggiornaCalcoli( self, event ):
        global contaPrivate

        self.grigliaLezioniSingole.ClearGrid()
        for i in range(0, 31, 1):
            #self.griglia.SetCellBackgroundColour(0, i, wx.GREEN)
            self.grigliaLezioniSingole.SetCellBackgroundColour(i, 0,wx.WHITE)
            self.grigliaLezioniSingole.SetCellBackgroundColour(i, 1, wx.WHITE)
            self.grigliaLezioniSingole.SetCellBackgroundColour(i, 2, wx.WHITE)
            self.grigliaLezioniSingole.SetCellBackgroundColour(i, 3, wx.WHITE)
            self.grigliaLezioniSingole.SetCellBackgroundColour(i, 4, wx.WHITE)
            self.grigliaLezioniSingole.SetCellBackgroundColour(i, 5, wx.WHITE)
        self.grigliaLezioniSingole.SetCellValue(31,0,'0')
        self.grigliaLezioniSingole.SetCellValue(31, 1, '0')
        self.grigliaLezioniSingole.SetCellValue(31, 2, '0')
        self.grigliaLezioniSingole.SetCellValue(31, 3, '0')
        self.grigliaLezioniSingole.SetCellValue(31, 4, '0')
        self.grigliaLezioniSingole.SetCellValue(31, 5, '0')
        self.grigliaTotali.SetCellValue(0, 0, '0')
        self.grigliaTotali.SetCellValue(0, 1, '0')
        self.grigliaTotali.SetCellValue(0, 2, '0')
        self.grigliaTotali.SetCellValue(0, 3, '0')
        self.grigliaTotali.SetCellValue(0, 4, '0')
        self.grigliaTotali.SetCellValue(0, 5, '0')
        self.grigliaTotali.SetCellValue(1, 0, '0')
        self.grigliaTotali.SetCellValue(1, 1, '0')
        self.grigliaTotali.SetCellValue(1, 2, '0')
        self.grigliaTotali.SetCellValue(1, 3, '0')
        self.grigliaTotali.SetCellValue(1, 4, '0')
        self.grigliaTotali.SetCellValue(1, 5, '0')
        self.grigliaTotali.SetCellValue(2, 0, '0')
        self.grigliaTotali.SetCellValue(2, 1, '0')
        self.grigliaTotali.SetCellValue(2, 2, '0')
        self.grigliaTotali.SetCellValue(2, 3, '0')
        self.grigliaTotali.SetCellValue(2, 4, '0')
        self.grigliaTotali.SetCellValue(2, 5, '0')

        datavecchia = str(self.calendarioStudenti.Date)
        nomeFile = datavecchia
        lezioniPrivate = 0

        nomeFile = nomeFile.replace('/', '-')
        nomeFile = nomeFile.replace(' 00:00:00', '')

        anno = '20' + nomeFile[-2:]
        percorso = './csv/' + anno + '/' + nomeFile[:2] + '/'
        if not os.path.exists(percorso):
            os.makedirs(percorso)
        print  percorso
        files = os.listdir(percorso)


        files_txt = [i for i in files if i.endswith('.csv')]
        # print files_txt
        # contaPrivate = 0
        # for files in files_txt:
        #     self.riempiTabella(percorso, files)

        # print files_txt
        print  files_txt

        popolaCalcoli = tabellaCalcoli.find_one(name=self.listaStudenti.StringSelection, anno=self.calendarioStudenti.Date.Year, mese=self.calendarioStudenti.Date.Month)
        popolaCalcoliMesePassato = tabellaCalcoli.find_one(name=self.listaStudenti.StringSelection,
                                                anno=self.calendarioStudenti.Date.Year,
                                                mese=self.calendarioStudenti.Date.Month-1)
        print  self.calendarioStudenti.Date.Month, 'self.calendarioStudenti.Date.Month'
        if popolaCalcoli is not  None:
            self.grigliaTotali.SetCellValue(0,0,popolaCalcoli.normaleigo)
            self.grigliaTotali.SetCellValue(0, 1, popolaCalcoli.normalsuugaku)
            self.grigliaTotali.SetCellValue(0, 2, popolaCalcoli.normalkokugo)
            self.grigliaTotali.SetCellValue(0, 3, popolaCalcoli.normalrika)
            self.grigliaTotali.SetCellValue(0, 4, popolaCalcoli.normalshakai)
            self.grigliaTotali.SetCellValue(0, 5, popolaCalcoli.normaltokubetsu)
            # self.grigliaTotali.SetCellValue(2, 0, popolaCalcoli.tsuikaeigo)
            # self.grigliaTotali.SetCellValue(2, 1, popolaCalcoli.tsuikasuugaku)
            # self.grigliaTotali.SetCellValue(2, 2, popolaCalcoli.tsuikakokugo)
            # self.grigliaTotali.SetCellValue(2, 3, popolaCalcoli.tsuikarika)
            # self.grigliaTotali.SetCellValue(2, 4, popolaCalcoli.tsuikashakai)
            # self.grigliaTotali.SetCellValue(2, 5, popolaCalcoli.tsuikatokubetsu)
        if popolaCalcoliMesePassato is not  None and popolaCalcoliMesePassato.balanceeigo is not None:
            self.grigliaTotali.SetCellValue(1, 0, popolaCalcoliMesePassato.balanceeigo)
            self.grigliaTotali.SetCellValue(1, 1, popolaCalcoliMesePassato.balancesuugaku)
            self.grigliaTotali.SetCellValue(1, 2, popolaCalcoliMesePassato.balancekokugo)
            self.grigliaTotali.SetCellValue(1, 3, popolaCalcoliMesePassato.balancerika)
            self.grigliaTotali.SetCellValue(1, 4, popolaCalcoliMesePassato.balanceshakai)
            self.grigliaTotali.SetCellValue(1, 5, popolaCalcoliMesePassato.balancetokubetu)
        if popolaCalcoliMesePassato is  None:
            self.grigliaTotali.SetCellValue(1, 0, '0')
            self.grigliaTotali.SetCellValue(1, 1, '0')
            self.grigliaTotali.SetCellValue(1, 2, '0')
            self.grigliaTotali.SetCellValue(1, 3, '0')
            self.grigliaTotali.SetCellValue(1, 4, '0')
            self.grigliaTotali.SetCellValue(1, 5, '0')
        if popolaCalcoli is None:
            self.grigliaTotali.SetCellValue(0,0,'0')
            self.grigliaTotali.SetCellValue(0, 1, '0')
            self.grigliaTotali.SetCellValue(0, 2,'0')
            self.grigliaTotali.SetCellValue(0, 3, '0')
            self.grigliaTotali.SetCellValue(0, 4,'0')
            self.grigliaTotali.SetCellValue(0, 5, '0')
            # self.grigliaTotali.SetCellValue(2, 0, '0')
            # self.grigliaTotali.SetCellValue(2, 1, '0')
            # self.grigliaTotali.SetCellValue(2, 2,'0')
            # self.grigliaTotali.SetCellValue(2, 3, '0')
            # self.grigliaTotali.SetCellValue(2, 4, '0')
            # self.grigliaTotali.SetCellValue(2, 5, '0')

        if files_txt is not None:
            contaPrivate = 0
            for files in files_txt:
                self.riempiTabella(percorso, files)
        if files_txt is None:
            self.grigliaLezioniSingole.ClearGrid()


        if popolaCalcoli is not None:
            datiCalcoli = (
            dict(id=popolaCalcoli.id, name=self.casellaNome.Value, mese=self.calendarioStudenti.Date.Month,
                 anno=self.calendarioStudenti.Date.Year, normaleigo=self.grigliaTotali.GetCellValue(0, 0),
                 normalsuugaku=self.grigliaTotali.GetCellValue(0, 1),
                 normalkokugo=self.grigliaTotali.GetCellValue(0, 2),
                 normalrika=self.grigliaTotali.GetCellValue(0, 3), normalshakai=self.grigliaTotali.GetCellValue(0, 4),
                 normaltokubetsu=self.grigliaTotali.GetCellValue(0, 5),
                 # tsuikaeigo=self.grigliaTotali.GetCellValue(2, 0),
                 # tsuikasuugaku=self.grigliaTotali.GetCellValue(2, 1),
                 # tsuikakokugo=self.grigliaTotali.GetCellValue(2, 2),
                 # tsuikarika=self.grigliaTotali.GetCellValue(2, 3),
                 # tsuikashakai=self.grigliaTotali.GetCellValue(2, 4),
                 # tsuikatokubetsu=self.grigliaTotali.GetCellValue(2, 5, ),
                #balanceeigo = self.grigliaLezioniSingole.GetCellValue(33, 0),
                #balancesuugaku = self.grigliaLezioniSingole.GetCellValue(33, 1),
                #balancekokugo = self.grigliaLezioniSingole.GetCellValue(33, 2),

                #balancerika = self.grigliaLezioniSingole.GetCellValue(33, 3),
                #balanceshakai = self.grigliaLezioniSingole.GetCellValue(33, 4),
                #balancetokubetu = self.grigliaLezioniSingole.GetCellValue(33, 5)
                 ))
            tabellaCalcoli.update(datiCalcoli, ['id'])
        if popolaCalcoli is None:
            tabellaCalcoli.insert(dict(name=self.casellaNome.Value, anno=self.calendarioStudenti.Date.Year,
                                       mese=self.calendarioStudenti.Date.Month,
                                       normaleigo=self.grigliaTotali.GetCellValue(0, 0),
                                       normalsuugaku=self.grigliaTotali.GetCellValue(0, 1),
                                       normalkokugo=self.grigliaTotali.GetCellValue(0, 2),
                                       normalrika=self.grigliaTotali.GetCellValue(0, 3),
                                       normalshakai=self.grigliaTotali.GetCellValue(0, 4),
                                       normaltokubetsu=self.grigliaTotali.GetCellValue(0, 5),
                                       # tsuikaeigo=self.grigliaTotali.GetCellValue(2, 0),
                                       # tsuikasuugaku=self.grigliaTotali.GetCellValue(2, 1),
                                       # tsuikakokugo=self.grigliaTotali.GetCellValue(2, 2),
                                       # tsuikarika=self.grigliaTotali.GetCellValue(2, 3),
                                       # tsuikashakai=self.grigliaTotali.GetCellValue(2, 4),
                                       # tsuikatokubetsu=self.grigliaTotali.GetCellValue(2, 5),
                                       # balanceeigo= self.grigliaLezioniSingole.GetCellValue(33,0),
                                       # balancesuugaku = self.grigliaLezioniSingole.GetCellValue(33,1),
                                       # balancekokugo=self.grigliaLezioniSingole.GetCellValue(33, 2),
                                       # balancerika=self.grigliaLezioniSingole.GetCellValue(33, 3),
                                       # balanceshakai=self.grigliaLezioniSingole.GetCellValue(33, 4),
                                       # balancetokubetu=self.grigliaLezioniSingole.GetCellValue(33, 5)
                                       ))
        totaleeigo = int(self.grigliaTotali.GetCellValue(0, 0))+int(self.grigliaTotali.GetCellValue(1, 0))+int(self.grigliaTotali.GetCellValue(2, 0))
        totalesuugaku   = int(self.grigliaTotali.GetCellValue(0, 1)) + int(self.grigliaTotali.GetCellValue(1, 1))+ int(self.grigliaTotali.GetCellValue(2, 1))
        totalekokugo = int(self.grigliaTotali.GetCellValue(0, 2))+int(self.grigliaTotali.GetCellValue(1, 2))+int(self.grigliaTotali.GetCellValue(2, 2))
        totalerika   = int(self.grigliaTotali.GetCellValue(0, 3)) + int(self.grigliaTotali.GetCellValue(1, 3))+ int(self.grigliaTotali.GetCellValue(2,3))
        totaleshakai = int(self.grigliaTotali.GetCellValue(0, 4))+int(self.grigliaTotali.GetCellValue(1, 4))+int(self.grigliaTotali.GetCellValue(2,4))
        totaletokubetsu   = int(self.grigliaTotali.GetCellValue(0, 5)) + int(self.grigliaTotali.GetCellValue(1, 5))+ int(self.grigliaTotali.GetCellValue(2, 5))
        self.grigliaTotali.SetCellValue(3, 0,str(totaleeigo))
        self.grigliaTotali.SetCellValue(3, 1, str(totalesuugaku))
        self.grigliaTotali.SetCellValue(3, 2, str(totalekokugo))
        self.grigliaTotali.SetCellValue(3, 3, str(totalerika))
        self.grigliaTotali.SetCellValue(3, 4, str(totaleshakai))
        self.grigliaTotali.SetCellValue(3, 5, str(totaletokubetsu))
        nokorieigo = int(self.grigliaTotali.GetCellValue(3, 0)) - int(self.grigliaLezioniSingole.GetCellValue(31, 0))
        nokorisuugaku = int(self.grigliaTotali.GetCellValue(3, 1)) - int(self.grigliaLezioniSingole.GetCellValue(31, 1))
        nokorikokugo = int(self.grigliaTotali.GetCellValue(3, 2)) - int(self.grigliaLezioniSingole.GetCellValue(31, 2))
        nokoririka = int(self.grigliaTotali.GetCellValue(3, 3)) - int(self.grigliaLezioniSingole.GetCellValue(31, 3))
        nokorishakai = int(self.grigliaTotali.GetCellValue(3, 4)) - int(self.grigliaLezioniSingole.GetCellValue(31,4))
        nokoritokubetsu = int(self.grigliaTotali.GetCellValue(3, 5)) - int(self.grigliaLezioniSingole.GetCellValue(31, 5))

        # self.grigliaLezioniSingole.SetCellValue(32,0,str(nokorieigo))
        # self.grigliaLezioniSingole.SetCellValue(32, 1, str(nokorisuugaku))
        # self.grigliaLezioniSingole.SetCellValue(32, 2, str(nokorikokugo))
        # self.grigliaLezioniSingole.SetCellValue(32, 3, str(nokoririka))
        # self.grigliaLezioniSingole.SetCellValue(32, 4, str(nokorishakai))
        # self.grigliaLezioniSingole.SetCellValue(32, 5, str(nokoritokubetsu))
    def selezionaCellaMateria1( self, event ):
        global rigaMaterie1
        global colonnaMaterie1
        rigaMaterie1 = event.GetRow()
        colonnaMaterie1 = event.GetCol()
        self.oreMaterie1.SetCellValue(rigaMaterie1, colonnaMaterie1, self.materieGiorni1.StringSelection)
    def caricaDate(self, event):
        global materieArray
        global materieTesto
        global contaPrivate
        contaPrivate = 0
        self.aggiorna.Enabled = True
        self.cancella.Enabled = True
        self.nuovo.Enabled = True
        self.materieGiorni.Clear()
        self.materieGiorni1.Clear()
        self.grigliaLezioniSingole.ClearGrid()
        self.grigliaTotali.ClearGrid()
        self.errorCheck.LabelText='-------------------------------------------------------------------------------------------------------------------------------------------------'
        self.usciteStudenti.Clear()
        self.invio.Enabled= False
        self.oreMaterie.ClearGrid()
        self.oreMaterie1.ClearGrid()
        for i in range(0, 31, 1):
            #self.griglia.SetCellBackgroundColour(0, i, wx.GREEN)
            self.grigliaLezioniSingole.SetCellBackgroundColour(i, 0,wx.WHITE)
            self.grigliaLezioniSingole.SetCellBackgroundColour(i, 1, wx.WHITE)
            self.grigliaLezioniSingole.SetCellBackgroundColour(i, 2, wx.WHITE)
            self.grigliaLezioniSingole.SetCellBackgroundColour(i, 3, wx.WHITE)
            self.grigliaLezioniSingole.SetCellBackgroundColour(i, 4, wx.WHITE)
            self.grigliaLezioniSingole.SetCellBackgroundColour(i, 5, wx.WHITE)

        for i in range(0, 9, 1):
            #self.griglia.SetCellBackgroundColour(0, i, wx.GREEN)
            self.oreMaterie1.SetCellBackgroundColour(0, i,wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(1, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(2, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(3, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(4, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(5, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(6, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(7, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(8, i, wx.WHITE)
        for i in range(0, 9, 1):
            #self.griglia.SetCellBackgroundColour(0, i, wx.GREEN)
            self.oreMaterie.SetCellBackgroundColour(0, i,wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(1, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(2, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(3, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(4, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(5, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(6, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(7, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(8, i, wx.WHITE)
        popolaGiorni = tabellaGiorni.find(name=self.listaStudenti.StringSelection)
        popolaDate = tabellaTempo.find(name=self.listaStudenti.StringSelection)
        popolastudenti = tabella.find(name=self.listaStudenti.StringSelection, student='1')


        global idSelezionato
        global idGiorni
        global idCalcoli
        materieTesto = []
        popolaCalcoli = tabellaCalcoli.find_one(name=self.listaStudenti.StringSelection,
                                                anno=self.calendarioStudenti.Date.Year,
                                                mese=self.calendarioStudenti.Date.Month)
        if popolaCalcoli is not None:
            self.grigliaTotali.SetCellValue(0, 0, popolaCalcoli.normaleigo)
            self.grigliaTotali.SetCellValue(0, 1, popolaCalcoli.normalsuugaku)
            self.grigliaTotali.SetCellValue(0, 2, popolaCalcoli.normalkokugo)
            self.grigliaTotali.SetCellValue(0, 3, popolaCalcoli.normalrika)
            self.grigliaTotali.SetCellValue(0, 4, popolaCalcoli.normalshakai)
            self.grigliaTotali.SetCellValue(0, 5, popolaCalcoli.normaltokubetsu)
            # self.grigliaTotali.SetCellValue(2, 0, popolaCalcoli.tsuikaeigo)
            # self.grigliaTotali.SetCellValue(2, 1, popolaCalcoli.tsuikasuugaku)
            # self.grigliaTotali.SetCellValue(2, 2, popolaCalcoli.tsuikakokugo)
            # self.grigliaTotali.SetCellValue(2, 3, popolaCalcoli.tsuikarika)
            # self.grigliaTotali.SetCellValue(2, 4, popolaCalcoli.tsuikashakai)
            # self.grigliaTotali.SetCellValue(2, 5, popolaCalcoli.tsuikatokubetsu)
        if popolaCalcoli is None:
            self.grigliaTotali.SetCellValue(0, 0, '0')
            self.grigliaTotali.SetCellValue(0, 1, '0')
            self.grigliaTotali.SetCellValue(0, 2, '0')
            self.grigliaTotali.SetCellValue(0, 3, '0')
            self.grigliaTotali.SetCellValue(0, 4, '0')
            self.grigliaTotali.SetCellValue(0, 5, '0')
            # self.grigliaTotali.SetCellValue(2, 0, '0')
            # self.grigliaTotali.SetCellValue(2, 1, '0')
            # self.grigliaTotali.SetCellValue(2, 2, '0')
            # self.grigliaTotali.SetCellValue(2, 3, '0')
            # self.grigliaTotali.SetCellValue(2, 4, '0')
            # self.grigliaTotali.SetCellValue(2, 5, '0')
        popolaCalcoliMesePassato = tabellaCalcoli.find_one(name=self.listaStudenti.StringSelection,
                                                           anno=self.calendarioStudenti.Date.Year,
                                                           mese=self.calendarioStudenti.Date.Month - 1)
        if popolaCalcoliMesePassato is not  None and popolaCalcoliMesePassato.balanceeigo is not None:
            self.grigliaTotali.SetCellValue(1, 0, popolaCalcoliMesePassato.balanceeigo)
            self.grigliaTotali.SetCellValue(1, 1, popolaCalcoliMesePassato.balancesuugaku)
            self.grigliaTotali.SetCellValue(1, 2, popolaCalcoliMesePassato.balancekokugo)
            self.grigliaTotali.SetCellValue(1, 3, popolaCalcoliMesePassato.balancerika)
            self.grigliaTotali.SetCellValue(1, 4, popolaCalcoliMesePassato.balanceshakai)
            self.grigliaTotali.SetCellValue(1, 5, popolaCalcoliMesePassato.balancetokubetu)
        if popolaCalcoliMesePassato is  not None and popolaCalcoliMesePassato.balanceeigo is  None:
            self.grigliaTotali.SetCellValue(1, 0, '0')
            self.grigliaTotali.SetCellValue(1, 1, '0')
            self.grigliaTotali.SetCellValue(1, 2, '0')
            self.grigliaTotali.SetCellValue(1, 3, '0')
            self.grigliaTotali.SetCellValue(1, 4, '0')
            self.grigliaTotali.SetCellValue(1, 5, '0')
        if popolaCalcoliMesePassato is None:
            self.grigliaTotali.SetCellValue(1, 0, '0')
            self.grigliaTotali.SetCellValue(1, 1, '0')
            self.grigliaTotali.SetCellValue(1, 2, '0')
            self.grigliaTotali.SetCellValue(1, 3, '0')
            self.grigliaTotali.SetCellValue(1, 4, '0')
            self.grigliaTotali.SetCellValue(1, 5, '0')
            #self.grigliaTotali.SetCellValue(0,0,i.)
        for i in range(0, 9, 1):
            #self.griglia.SetCellBackgroundColour(0, i, wx.GREEN)
            self.oreMaterie.SetCellBackgroundColour(0, i,wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(1, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(2, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(3, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(4, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(5, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(6, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(7, i, wx.WHITE)
            self.oreMaterie.SetCellBackgroundColour(8, i, wx.WHITE)
        for i in range(0, 9, 1):
            #self.griglia.SetCellBackgroundColour(0, i, wx.GREEN)
            self.oreMaterie1.SetCellBackgroundColour(0, i,wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(1, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(2, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(3, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(4, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(5, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(6, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(7, i, wx.WHITE)
            self.oreMaterie1.SetCellBackgroundColour(8, i, wx.WHITE)
        for i in popolastudenti:

            idSelezionato = i.id
            self.furigana.Value = i.furigana
            self.joseiBox.Value =i.femmina
            self.danseiBox.Value = i.maschio
            self.sonota.Value = i.sonota
            self.casellaScuola.LabelText = i.scuola
            self.casellaNome.LabelText = i.name
            self.casellaEmail.LabelText = i.email
            self.mailGenitori.LabelText = i.parentMail
            self.casellaTelefono.LabelText = i.telephone
            self.eigo.Value = i.eigo
            self.rika.Value = i.rika
            self.shakai.Value = i.shakai
            self.suugaku.Value = i.suugaku
            self.tokubetsu.Value = i.tokubetsu
            self.kokugo.Value = i.kokugo
            self.eigo1.Value = i.eigo1
            self.rika1.Value = i.rika1
            self.shakai1.Value = i.shakai1
            self.suugaku1.Value = i.suugaku1
            self.tokubetsu1.Value = i.tokubetsu1
            self.kokugo1.Value = i.kokugo1
            self.individual.Value = i.individual
            self.shared.Value = i.shared

            materieArray=[i.eigo,i.rika,i.shakai,i.suugaku,i.tokubetsu,i.kokugo]
            materieArray1 = [i.eigo, i.rika, i.shakai, i.suugaku, i.tokubetsu, i.kokugo]

            if self.eigo.Value == True:
                self.materieGiorni.Append(u'英語')
            if self.rika.Value == True:
                self.materieGiorni.Append(u'理科')
            if self.shakai.Value == True:
                self.materieGiorni.Append(u'社会')
            if self.suugaku.Value == True:
                self.materieGiorni.Append(u'数学')
            if self.tokubetsu.Value == True:
                self.materieGiorni.Append(u'特別')
            if self.kokugo.Value == True:
                self.materieGiorni.Append(u'国語')

            if self.eigo1.Value == True:
                self.materieGiorni1.Append(u'英語')
            if self.rika1.Value == True:
                self.materieGiorni1.Append(u'理科')
            if self.shakai1.Value == True:
                self.materieGiorni1.Append(u'社会')
            if self.suugaku1.Value == True:
                self.materieGiorni1.Append(u'数学')
            if self.tokubetsu1.Value == True:
                self.materieGiorni1.Append(u'特別')
            if self.kokugo1.Value == True:
                self.materieGiorni1.Append(u'国語')

            if i.cardID == '' or i.cardID == u"カード未登録です、登録してください" :
                self.cardid.LabelText = u"カード未登録です、登録してください"
                self.cardcancel.Enabled = False
                self.CardRegistration.Enabled= True

            else:
                self.cardid.LabelText = i.cardID
                self.cardcancel.Enabled = True
                self.CardRegistration.Enabled= False
            arrayore = [i.primaOra, i.secondaOra, i.terzaOra, i.quartaOra, i.quintaOra, i.sestaOra, i.settimaOra,
                        i.ottavaOra, i.nonaOra]
            arrayore1 = [i.primaOra1, i.secondaOra1, i.terzaOra1, i.quartaOra1, i.quintaOra1, i.sestaOra1, i.settimaOra1,
                        i.ottavaOra1, i.nonaOra1]
            #print arrayore
            for settaOre in range(0, 9, 1):
                self.tabellaOre.SetSelection(settaOre, select=arrayore[settaOre])
            for settaOre in range(0, 9, 1):
                self.tabellaOre1.SetSelection(settaOre, select=arrayore1[settaOre])


        for i in popolaDate:
            if len((str(i.uscite)))>=5:
                self.usciteStudenti.Append(str(i.uscite))


        for i in popolaGiorni:
            idGiorni  = i.id
            self.lunedi.Value = i.lunedi
            self.martedi.Value = i.martedi
            self.mercoledi.Value = i.mercoledi
            self.giovedi.Value = i.giovedi
            self.venerdi.Value = i.venerdi
            self.sabato.Value = i.sabato
            self.domenica.Value = i.domenica
            self.lunedi1.Value = i.lunedi1
            self.martedi1.Value = i.martedi1
            self.mercoledi1.Value = i.mercoledi1
            self.giovedi1.Value = i.giovedi1
            self.venerdi1.Value = i.venerdi1
            self.sabato1.Value = i.sabato1
            self.domenica1.Value = i.domenica1

        datavecchia = str(self.calendarioStudenti.Date)
        nomeFile = datavecchia
        lezioniPrivate = 0

        nomeFile = nomeFile.replace('/', '-')
        nomeFile = nomeFile.replace(' 00:00:00', '')

        anno = '20' + nomeFile[-2:]
        percorso = './csv/' + anno + '/' + nomeFile[:2] + '/'
        listashift = []
        listashift1 = []
        if self.lunedi.Value == True:
            listashift.append(0)
        if self.martedi.Value == True:
            listashift.append(1)
        if self.mercoledi.Value == True:
            listashift.append(2)
        if self.giovedi.Value == True:
            listashift.append(3)
        if self.venerdi.Value == True:
            listashift.append(4)
        if self.sabato.Value == True:
            listashift.append(5)
        if self.domenica.Value == True:
            listashift.append(6)

        if self.lunedi1.Value == True:
            listashift1.append(0)
        if self.martedi1.Value == True:
            listashift1.append(1)
        if self.mercoledi1.Value == True:
            listashift1.append(2)
        if self.giovedi1.Value == True:
            listashift1.append(3)
        if self.venerdi1.Value == True:
            listashift1.append(4)
        if self.sabato1.Value == True:
            listashift1.append(5)
        if self.domenica1.Value == True:
            listashift1.append(6)
        self.grigliaLezioniSingole.SetColFormatNumber(0)
        self.grigliaLezioniSingole.SetColFormatNumber(1)
        self.grigliaLezioniSingole.SetColFormatNumber(2)
        self.grigliaLezioniSingole.SetColFormatNumber(3)
        self.grigliaLezioniSingole.SetColFormatNumber(4)
        self.grigliaLezioniSingole.SetColFormatNumber(5)
        contagiri = 0
        contagiri1 = 0
        lunghezzaShift = len(listashift)
        lunghezzaShift1 = len(listashift1)

        for i in range(0,9,1):
            if lunghezzaShift >=1:
                for giorni in listashift:
                    if arrayore[contagiri] == True:
                        self.oreMaterie.SetCellBackgroundColour(contagiri, giorni, wx.GREEN)
            contagiri = contagiri+1
        self.oreMaterie.Refresh()
        # percorsoStudentimemo = './StudentsData/' + self.casellaNome.Value + self.casellaTelefono.Value + 'memo.txt'
        # controllaPercorso = os.path.exists(percorsoStudentimemo)
        #
        # if controllaPercorso == True:
        #     with open(percorsoStudentimemo, 'rb') as f:
        #         self.memo.Value = f
        percorsoStudenti = './StudentsData/'+self.casellaNome.Value+self.casellaTelefono.Value+'.txt'
        controllaPercorso = os.path.exists(percorsoStudenti)

        if controllaPercorso == True:
            with open(percorsoStudenti, 'rb') as f:
                reader = csv.DictReader(f)
                contarighe = 0
                converti = csvkit.unicsv.UnicodeCSVDictReader(f=f, encoding='utf-8')

                for i in converti:
                    self.oreMaterie.SetCellValue(contarighe, 0, i[u'月曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 1, i[u'火曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 2, i[u'水曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 3, i[u'木曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 4, i[u'金曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 5, i[u'土曜日'])
                    self.oreMaterie.SetCellValue(contarighe, 6, i[u'日曜日'])

                    contarighe = contarighe + 1
        for i in range(0,9,1):
            if lunghezzaShift1 >=1:
                for giorni in listashift1:
                    if arrayore1[contagiri1] == True:
                        self.oreMaterie1.SetCellBackgroundColour(contagiri1, giorni, wx.RED)
            contagiri1 = contagiri1+1
        self.oreMaterie1.Refresh()

        percorsoStudenti1 = './StudentsData/'+self.casellaNome.Value+self.casellaTelefono.Value+'tokubetsu.txt'
        controllaPercorso1 = os.path.exists(percorsoStudenti1)

        if controllaPercorso1 == True:
            with open(percorsoStudenti1, 'rb') as f1:
                reader1 = csv.DictReader(f1)
                contarighe1 = 0
                converti1 = csvkit.unicsv.UnicodeCSVDictReader(f=f1, encoding='utf-8')

                for i in converti1:
                    self.oreMaterie1.SetCellValue(contarighe1, 0, i[u'月曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 1, i[u'火曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 2, i[u'水曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 3, i[u'木曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 4, i[u'金曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 5, i[u'土曜日'])
                    self.oreMaterie1.SetCellValue(contarighe1, 6, i[u'日曜日'])

                    contarighe1 = contarighe1 + 1

        files = os.listdir(percorso)

        files_txt = [i for i in files if i.endswith('.csv')]
        #print files_txt

        for files in files_txt:

            self.riempiTabella(percorso,files)

        totaleeigo = int(self.grigliaTotali.GetCellValue(0, 0)) + int(self.grigliaTotali.GetCellValue(1, 0)) + int(
            self.grigliaTotali.GetCellValue(2, 0))
        totalesuugaku = int(self.grigliaTotali.GetCellValue(0, 1)) + int(self.grigliaTotali.GetCellValue(1, 1)) + int(
            self.grigliaTotali.GetCellValue(2, 1))
        totalekokugo = int(self.grigliaTotali.GetCellValue(0, 2)) + int(self.grigliaTotali.GetCellValue(1, 2)) + int(
            self.grigliaTotali.GetCellValue(2, 2))
        totalerika = int(self.grigliaTotali.GetCellValue(0, 3)) + int(self.grigliaTotali.GetCellValue(1, 3)) + int(
            self.grigliaTotali.GetCellValue(2, 3))
        totaleshakai = int(self.grigliaTotali.GetCellValue(0, 4)) + int(self.grigliaTotali.GetCellValue(1, 4)) + int(
            self.grigliaTotali.GetCellValue(2, 4))
        totaletokubetsu = int(self.grigliaTotali.GetCellValue(0, 5)) + int(self.grigliaTotali.GetCellValue(1, 5)) + int(
            self.grigliaTotali.GetCellValue(2, 5))

        self.grigliaTotali.SetCellValue(3, 0, str(totaleeigo))
        self.grigliaTotali.SetCellValue(3, 1, str(totalesuugaku))
        self.grigliaTotali.SetCellValue(3, 2, str(totalekokugo))
        self.grigliaTotali.SetCellValue(3, 3, str(totalerika))
        self.grigliaTotali.SetCellValue(3, 4, str(totaleshakai))
        self.grigliaTotali.SetCellValue(3, 5, str(totaletokubetsu))
        if totaleeigo == int(self.grigliaLezioniSingole.GetCellValue(34,0)):
            self.grigliaTotali.SetCellBackgroundColour(3,0,wx.GREEN)
        elif totaleeigo < int(self.grigliaLezioniSingole.GetCellValue(34,0)):
            self.grigliaTotali.SetCellBackgroundColour(3, 0, wx.RED)
        elif totaleeigo > int(self.grigliaLezioniSingole.GetCellValue(34,0)):
            self.grigliaTotali.SetCellBackgroundColour(3, 0, wx.YELLOW)
        if totalesuugaku == int(self.grigliaLezioniSingole.GetCellValue(34,1)):
            self.grigliaTotali.SetCellBackgroundColour(3,1,wx.GREEN)
        elif totalesuugaku < int(self.grigliaLezioniSingole.GetCellValue(34,1)):
            self.grigliaTotali.SetCellBackgroundColour(3, 1, wx.RED)
        elif totalesuugaku > int(self.grigliaLezioniSingole.GetCellValue(34,1)):
            self.grigliaTotali.SetCellBackgroundColour(3, 1,wx.YELLOW)

        if totalekokugo == int(self.grigliaLezioniSingole.GetCellValue(34, 2)):
            self.grigliaTotali.SetCellBackgroundColour(3, 2, wx.GREEN)
        elif totalekokugo < int(self.grigliaLezioniSingole.GetCellValue(34, 2)):
            self.grigliaTotali.SetCellBackgroundColour(3, 2, wx.RED)
        elif totalekokugo > int(self.grigliaLezioniSingole.GetCellValue(34, 2)):
            self.grigliaTotali.SetCellBackgroundColour(3, 2, wx.YELLOW)
        if totalerika == int(self.grigliaLezioniSingole.GetCellValue(34, 3)):
            self.grigliaTotali.SetCellBackgroundColour(3, 3, wx.GREEN)
        elif totalerika < int(self.grigliaLezioniSingole.GetCellValue(34, 3)):
            self.grigliaTotali.SetCellBackgroundColour(3, 3, wx.RED)
        elif totalerika > int(self.grigliaLezioniSingole.GetCellValue(34, 3)):
            self.grigliaTotali.SetCellBackgroundColour(3, 3, wx.YELLOW)

        if totaleshakai == int(self.grigliaLezioniSingole.GetCellValue(34, 4)):
            self.grigliaTotali.SetCellBackgroundColour(3, 4, wx.GREEN)
        elif totaleshakai < int(self.grigliaLezioniSingole.GetCellValue(34, 4)):
            self.grigliaTotali.SetCellBackgroundColour(3, 4, wx.RED)
        elif totaleshakai > int(self.grigliaLezioniSingole.GetCellValue(34, 4)):
            self.grigliaTotali.SetCellBackgroundColour(3, 4, wx.YELLOW)
        if totaletokubetsu == int(self.grigliaLezioniSingole.GetCellValue(34, 5)):
            self.grigliaTotali.SetCellBackgroundColour(3, 5, wx.GREEN)
        elif totaletokubetsu < int(self.grigliaLezioniSingole.GetCellValue(34, 5)):
            self.grigliaTotali.SetCellBackgroundColour(3, 5, wx.RED)
        elif totaletokubetsu > int(self.grigliaLezioniSingole.GetCellValue(34, 5)):
            self.grigliaTotali.SetCellBackgroundColour(3, 5, wx.YELLOW)
        # nokorieigo = int(self.grigliaTotali.GetCellValue(3, 0)) - int(self.grigliaLezioniSingole.GetCellValue(31, 0))
        # nokorisuugaku = int(self.grigliaTotali.GetCellValue(3, 1)) - int(self.grigliaLezioniSingole.GetCellValue(31, 1))
        # nokorikokugo = int(self.grigliaTotali.GetCellValue(3, 2)) - int(self.grigliaLezioniSingole.GetCellValue(31, 2))
        # nokoririka = int(self.grigliaTotali.GetCellValue(3, 3)) - int(self.grigliaLezioniSingole.GetCellValue(31, 3))
        # nokorishakai = int(self.grigliaTotali.GetCellValue(3, 4)) - int(self.grigliaLezioniSingole.GetCellValue(31, 4))
        nokorieigo = int(self.grigliaLezioniSingole.GetCellValue(31, 0))-int(self.grigliaTotali.GetCellValue(2, 0))
        nokorisuugaku = int(self.grigliaLezioniSingole.GetCellValue(31, 1))-int(self.grigliaTotali.GetCellValue(2, 1))
        nokorikokugo = int(self.grigliaLezioniSingole.GetCellValue(31, 2))-int(self.grigliaTotali.GetCellValue(2, 2))
        nokoririka = int(self.grigliaLezioniSingole.GetCellValue(31, 3))-int(self.grigliaTotali.GetCellValue(2, 3))
        nokorishakai = int(self.grigliaLezioniSingole.GetCellValue(31, 4))-int(self.grigliaTotali.GetCellValue(2, 4))
        nokoritokubetsu =  int(self.grigliaLezioniSingole.GetCellValue(31, 5))-int(self.grigliaTotali.GetCellValue(2, 5))
        # if nokorieigo < 0:
        #     self.grigliaTotali.SetCellValue(2,0,str(nokorieigo))
        # if nokorisuugaku < 0:
        #     self.grigliaTotali.SetCellValue(2,1,str(nokorisuugaku))
        # if nokorikokugo < 0:
        #     self.grigliaTotali.SetCellValue(2,2,str(nokorikokugo))
        # if nokoririka < 0:
        #     self.grigliaTotali.SetCellValue(2,3,str(nokoririka))
        # if nokorishakai < 0:
        #     self.grigliaTotali.SetCellValue(2,4,str(nokorishakai))
        # if nokoritokubetsu < 0:
        #     self.grigliaTotali.SetCellValue(2,5,str(nokoritokubetsu))
        #if nokorieigo < 0:
        # self.grigliaTotali.SetCellValue(2, 0, str(nokorieigo))
        # #if nokorisuugaku < 0:
        # self.grigliaTotali.SetCellValue(2, 1, str(nokorisuugaku))
        # #if nokorikokugo < 0:
        # self.grigliaTotali.SetCellValue(2, 2, str(nokorikokugo))
        # #if nokoririka < 0:
        # self.grigliaTotali.SetCellValue(2, 3, str(nokoririka))
        # #if nokorishakai < 0:
        # self.grigliaTotali.SetCellValue(2, 4, str(nokorishakai))
        # #if nokoritokubetsu < 0:
        # self.grigliaTotali.SetCellValue(2, 5, str(nokoritokubetsu))
        # self.grigliaLezioniSingole.SetCellValue(32, 0, str(nokorieigo))
        # self.grigliaLezioniSingole.SetCellValue(32, 1, str(nokorisuugaku))
        # self.grigliaLezioniSingole.SetCellValue(32, 2, str(nokorikokugo))
        # self.grigliaLezioniSingole.SetCellValue(32, 3, str(nokoririka))
        # self.grigliaLezioniSingole.SetCellValue(32, 4, str(nokorishakai))
        # self.grigliaLezioniSingole.SetCellValue(32, 5, str(nokoritokubetsu))
        #for i in listashift:
        #    print i
        self.invio.Enabled = False
    def riempiTabella(self,percorso,files):
        global contaPrivate
        with open(percorso + files, 'rb') as f:

            sommainglese = 0
            sommakokugo = 0
            sommashakai = 0
            sommarika = 0
            sommatokubetsu = 0
            sommasuugaku = 0
            reader = csv.DictReader(f)
            contarighe = 0
            converti = csvkit.unicsv.UnicodeCSVDictReader(f=f, encoding='utf-8')
            # print converti, converti
            kokugotemp = 0
            suugakutemp = 0
            eigotemp = 0
            rikatemp = 0
            shakaitemp = 0
            tokubetsutemp = 0
            kokugotemp1 = 0
            suugakutemp1 = 0
            eigotemp1 = 0
            rikatemp1 = 0
            shakaitemp1 = 0
            tokubetsutemp1 = 0
            dataComposta = funzioni.trasformaNomefileInOra(f.name)
            controlloCheckIn = funzioni.controlloCheckIn(self.listaStudenti.StringSelection, tabellaTempo, dataComposta)
            for i in converti:

                if controlloCheckIn == 'NON':
                    if self.listaStudenti.StringSelection in i['9:10 - 10:20']:
                        aggiunngimateria = funzioni.contalezioni(i['9:10 - 10:20'])
                        if aggiunngimateria == 1:
                            kokugotemp1 = kokugotemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2,
                                                                    str(kokugotemp1))
                        if aggiunngimateria == 2:
                            eigotemp1 = eigotemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0,
                                                                    str(eigotemp1))
                        if aggiunngimateria == 3:
                            suugakutemp1 = suugakutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1,
                                                                    str(suugakutemp1))
                        if aggiunngimateria == 4:
                            rikatemp1 = rikatemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3,
                                                                    str(rikatemp1))
                        if aggiunngimateria == 5:
                            shakaitemp1 = shakaitemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4,
                                                                    str(shakaitemp1))
                        if aggiunngimateria == 6:
                            tokubetsutemp1 = tokubetsutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5,
                                                                    str(tokubetsutemp1))
                    if self.listaStudenti.StringSelection in i['10:30 - 11:40']:
                        aggiunngimateria = funzioni.contalezioni(i['10:30 - 11:40'])
                        if aggiunngimateria == 1:
                            kokugotemp1 = kokugotemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2,
                                                                    str(kokugotemp1))
                        if aggiunngimateria == 2:
                            eigotemp1 = eigotemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0,
                                                                    str(eigotemp1))
                        if aggiunngimateria == 3:
                            suugakutemp1 = suugakutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1,
                                                                    str(suugakutemp1))
                        if aggiunngimateria == 4:
                            rikatemp1 = rikatemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3,
                                                                    str(rikatemp1))
                        if aggiunngimateria == 5:
                            shakaitemp1 = shakaitemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4,
                                                                    str(shakaitemp1))
                        if aggiunngimateria == 6:
                            tokubetsutemp1 = tokubetsutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5,
                                                                    str(tokubetsutemp1))
                    if self.listaStudenti.StringSelection in i['11:50 - 13:00']:
                        aggiunngimateria = funzioni.contalezioni(i['11:50 - 13:00'])
                        if aggiunngimateria == 1:
                            kokugotemp1 = kokugotemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2,
                                                                    str(kokugotemp1))
                        if aggiunngimateria == 2:
                            eigotemp1 = eigotemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0,
                                                                    str(eigotemp1))
                        if aggiunngimateria == 3:
                            suugakutemp1 = suugakutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1,
                                                                    str(suugakutemp1))
                        if aggiunngimateria == 4:
                            rikatemp1 = rikatemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3,
                                                                    str(rikatemp1))
                        if aggiunngimateria == 5:
                            shakaitemp1 = shakaitemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4,
                                                                    str(shakaitemp1))
                        if aggiunngimateria == 6:
                            tokubetsutemp1 = tokubetsutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5,
                                                                    str(tokubetsutemp1))
                    if self.listaStudenti.StringSelection in i['13:40 - 14:50']:
                        aggiunngimateria = funzioni.contalezioni(i['13:40 - 14:50'])
                        if aggiunngimateria == 1:
                            kokugotemp1 = kokugotemp1 + 1

                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2,
                                                                    str(kokugotemp1))
                        if aggiunngimateria == 2:
                            eigotemp1 = eigotemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0,
                                                                    str(eigotemp1))
                        if aggiunngimateria == 3:
                            suugakutemp1 = suugakutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1,
                                                                    str(suugakutemp1))
                        if aggiunngimateria == 4:
                            rikatemp1 = rikatemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3,
                                                                    str(rikatemp1))
                        if aggiunngimateria == 5:
                            shakaitemp1 = shakaitemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4,
                                                                    str(shakaitemp1))
                        if aggiunngimateria == 6:
                            tokubetsutemp1 = tokubetsutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5,
                                                                    str(tokubetsutemp1))

                    if self.listaStudenti.StringSelection in i['15:00 - 16:10']:
                        aggiunngimateria = funzioni.contalezioni(i['15:00 - 16:10'])
                        if aggiunngimateria == 1:
                            kokugotemp1 = kokugotemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2,
                                                                    str(kokugotemp1))
                        if aggiunngimateria == 2:
                            eigotemp1 = eigotemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0,
                                                                    str(eigotemp1))
                        if aggiunngimateria == 3:
                            suugakutemp1 = suugakutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1,
                                                                    str(suugakutemp1))
                        if aggiunngimateria == 4:
                            rikatemp1 = rikatemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3,
                                                                    str(rikatemp1))
                        if aggiunngimateria == 5:
                            shakaitemp1 = shakaitemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4,
                                                                    str(shakaitemp1))
                        if aggiunngimateria == 6:
                            tokubetsutemp1 = tokubetsutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5,
                                                                    str(tokubetsutemp1))

                    if self.listaStudenti.StringSelection in i['16:40 - 17:50']:
                        aggiunngimateria = funzioni.contalezioni(i['16:40 - 17:50'])
                        if aggiunngimateria == 1:
                            kokugotemp1 = kokugotemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2,
                                                                    str(kokugotemp1))
                        if aggiunngimateria == 2:
                            eigotemp1 = eigotemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0,
                                                                    str(eigotemp1))
                        if aggiunngimateria == 3:
                            suugakutemp1 = suugakutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1,
                                                                    str(suugakutemp1))
                        if aggiunngimateria == 4:
                            rikatemp1 = rikatemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3,
                                                                    str(rikatemp1))
                        if aggiunngimateria == 5:
                            shakaitemp1 = shakaitemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4,
                                                                    str(shakaitemp1))
                        if aggiunngimateria == 6:
                            tokubetsutemp1 = tokubetsutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5,
                                                                    str(tokubetsutemp1))

                    if self.listaStudenti.StringSelection in i['18:00 - 19:10']:
                        aggiunngimateria = funzioni.contalezioni(i['18:00 - 19:10'])
                        if aggiunngimateria == 1:
                            kokugotemp1 = kokugotemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2,
                                                                    str(kokugotemp1))
                        if aggiunngimateria == 2:
                            eigotemp1 = eigotemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0,
                                                                    str(eigotemp1))
                        if aggiunngimateria == 3:
                            suugakutemp1 = suugakutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1,
                                                                    str(suugakutemp1))
                        if aggiunngimateria == 4:
                            rikatemp1 = rikatemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3,
                                                                    str(rikatemp1))
                        if aggiunngimateria == 5:
                            shakaitemp1 = shakaitemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4,
                                                                    str(shakaitemp1))
                        if aggiunngimateria == 6:
                            tokubetsutemp1 = tokubetsutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5,
                                                                    str(tokubetsutemp1))

                    if self.listaStudenti.StringSelection in i['19:20 - 20:30']:
                        aggiunngimateria = funzioni.contalezioni(i['19:20 - 20:30'])
                        if aggiunngimateria == 1:
                            kokugotemp1 = kokugotemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2,
                                                                    str(kokugotemp1))
                        if aggiunngimateria == 2:
                            eigotemp1 = eigotemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0,
                                                                    str(eigotemp1))
                        if aggiunngimateria == 3:
                            suugakutemp1= suugakutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1,
                                                                    str(suugakutemp1))
                        if aggiunngimateria == 4:
                            rikatemp1 = rikatemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3,
                                                                    str(rikatemp1))
                        if aggiunngimateria == 5:
                            shakaitemp1 = shakaitemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4,
                                                                    str(shakaitemp1))
                        if aggiunngimateria == 6:
                            tokubetsutemp1 = tokubetsutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5,
                                                                    str(tokubetsutemp1))

                    if self.listaStudenti.StringSelection in i['20:40 - 21:50']:
                        aggiunngimateria = funzioni.contalezioni(i['20:40 - 21:50'])
                        if aggiunngimateria == 1:
                            kokugotemp1 = kokugotemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2,
                                                                    str(kokugotemp1))
                        if aggiunngimateria == 2:
                            eigotemp1 = eigotemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0,
                                                                    str(eigotemp1))
                        if aggiunngimateria == 3:
                            suugakutemp1 = suugakutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1,
                                                                    str(suugakutemp1))
                        if aggiunngimateria == 4:
                            rikatemp1 = rikatemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3,
                                                                    str(rikatemp1))
                        if aggiunngimateria == 5:
                            shakaitemp1 = shakaitemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4,
                                                                    str(shakaitemp1))
                        if aggiunngimateria == 6:
                            tokubetsutemp1 = tokubetsutemp1 + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5,
                                                                    str(tokubetsutemp1))
                            # print  lezioniPrivate, 'lezioni private'
                if controlloCheckIn == 'OUT' or controlloCheckIn == 'IN':

                    if self.listaStudenti.StringSelection in i['9:10 - 10:20']:
                        aggiunngimateria = funzioni.contalezioni(i['9:10 - 10:20'])
                        if aggiunngimateria == 1:
                            kokugotemp = kokugotemp +1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2, str(kokugotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 2, wx.GREEN)
                        if aggiunngimateria == 2:
                            eigotemp = eigotemp+1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0, str(eigotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 0, wx.GREEN)
                        if aggiunngimateria == 3:
                            suugakutemp = suugakutemp +1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1, str(suugakutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 1, wx.GREEN)
                        if aggiunngimateria == 4:
                            rikatemp = rikatemp+1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3, str(rikatemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 3, wx.GREEN)
                        if aggiunngimateria == 5:
                            shakaitemp = shakaitemp+1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4, str(shakaitemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 4, wx.GREEN)
                        if aggiunngimateria == 6:
                            tokubetsutemp = tokubetsutemp+1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5, str(tokubetsutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 5, wx.GREEN)
                    if self.listaStudenti.StringSelection in i['10:30 - 11:40']:
                        aggiunngimateria = funzioni.contalezioni(i['10:30 - 11:40'])
                        if aggiunngimateria == 1:
                            kokugotemp = kokugotemp +1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2, str(kokugotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 2, wx.GREEN)
                        if aggiunngimateria == 2:
                            eigotemp = eigotemp+1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0, str(eigotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 0, wx.GREEN)
                        if aggiunngimateria == 3:
                            suugakutemp = suugakutemp +1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1, str(suugakutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 1, wx.GREEN)
                        if aggiunngimateria == 4:
                            rikatemp = rikatemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3, str(rikatemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 3, wx.GREEN)
                        if aggiunngimateria == 5:
                            shakaitemp = shakaitemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4, str(shakaitemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 4, wx.GREEN)
                        if aggiunngimateria == 6:
                            tokubetsutemp = tokubetsutemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5, str(tokubetsutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 5, wx.GREEN)
                    if self.listaStudenti.StringSelection in i['11:50 - 13:00']:
                        aggiunngimateria = funzioni.contalezioni(i['11:50 - 13:00'])
                        if aggiunngimateria == 1:
                            kokugotemp = kokugotemp +1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2, str(kokugotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 2, wx.GREEN)
                        if aggiunngimateria == 2:
                            eigotemp = eigotemp+1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0, str(eigotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 0, wx.GREEN)
                        if aggiunngimateria == 3:
                            suugakutemp = suugakutemp +1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1, str(suugakutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 1, wx.GREEN)
                        if aggiunngimateria == 4:
                            rikatemp = rikatemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3, str(rikatemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 3, wx.GREEN)
                        if aggiunngimateria == 5:
                            shakaitemp = shakaitemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4, str(shakaitemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 4, wx.GREEN)
                        if aggiunngimateria == 6:
                            tokubetsutemp = tokubetsutemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5, str(tokubetsutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 5, wx.GREEN)
                    if self.listaStudenti.StringSelection in i['13:40 - 14:50']:
                        aggiunngimateria = funzioni.contalezioni(i['13:40 - 14:50'])
                        if aggiunngimateria == 1:
                            kokugotemp = kokugotemp +1

                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2, str(kokugotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 2, wx.GREEN)
                        if aggiunngimateria == 2:
                            eigotemp = eigotemp+1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0, str(eigotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 0, wx.GREEN)
                        if aggiunngimateria == 3:
                            suugakutemp = suugakutemp +1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1, str(suugakutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 1, wx.GREEN)
                        if aggiunngimateria == 4:
                            rikatemp = rikatemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3, str(rikatemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 3, wx.GREEN)
                        if aggiunngimateria == 5:
                            shakaitemp = shakaitemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4, str(shakaitemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 4, wx.GREEN)
                        if aggiunngimateria == 6:
                            tokubetsutemp = tokubetsutemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5, str(tokubetsutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 5, wx.GREEN)

                    if self.listaStudenti.StringSelection in i['15:00 - 16:10']:
                        aggiunngimateria = funzioni.contalezioni(i['15:00 - 16:10'])
                        if aggiunngimateria == 1:
                            kokugotemp = kokugotemp +1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2, str(kokugotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 2, wx.GREEN)
                        if aggiunngimateria == 2:
                            eigotemp = eigotemp+1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0, str(eigotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 0, wx.GREEN)
                        if aggiunngimateria == 3:
                            suugakutemp = suugakutemp +1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1, str(suugakutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 1, wx.GREEN)
                        if aggiunngimateria == 4:
                            rikatemp = rikatemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3, str(rikatemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 3, wx.GREEN)
                        if aggiunngimateria == 5:
                            shakaitemp = shakaitemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4, str(shakaitemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 4, wx.GREEN)
                        if aggiunngimateria == 6:
                            tokubetsutemp = tokubetsutemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5, str(tokubetsutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 5, wx.GREEN)

                    if self.listaStudenti.StringSelection in i['16:40 - 17:50']:
                        aggiunngimateria = funzioni.contalezioni(i['16:40 - 17:50'])
                        if aggiunngimateria == 1:
                            kokugotemp = kokugotemp +1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2, str(kokugotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 2, wx.GREEN)
                        if aggiunngimateria == 2:
                            eigotemp = eigotemp+1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0, str(eigotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 0, wx.GREEN)
                        if aggiunngimateria == 3:
                            suugakutemp = suugakutemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1, str(suugakutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 1, wx.GREEN)
                        if aggiunngimateria == 4:
                            rikatemp = rikatemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3, str(rikatemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 3, wx.GREEN)
                        if aggiunngimateria == 5:
                            shakaitemp = shakaitemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4, str(shakaitemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 4, wx.GREEN)
                        if aggiunngimateria == 6:
                            tokubetsutemp = tokubetsutemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5, str(tokubetsutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 5, wx.GREEN)

                    if self.listaStudenti.StringSelection in i['18:00 - 19:10']:
                        aggiunngimateria = funzioni.contalezioni(i['18:00 - 19:10'])
                        if aggiunngimateria == 1:
                            kokugotemp = kokugotemp +1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2, str(kokugotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 2, wx.GREEN)
                        if aggiunngimateria == 2:
                            eigotemp = eigotemp+1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0, str(eigotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 0, wx.GREEN)
                        if aggiunngimateria == 3:
                            suugakutemp = suugakutemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1, str(suugakutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 1, wx.GREEN)
                        if aggiunngimateria == 4:
                            rikatemp = rikatemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3, str(rikatemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 3, wx.GREEN)
                        if aggiunngimateria == 5:
                            shakaitemp = shakaitemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4, str(shakaitemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 4, wx.GREEN)
                        if aggiunngimateria == 6:
                            tokubetsutemp = tokubetsutemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5, str(tokubetsutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 5, wx.GREEN)

                    if self.listaStudenti.StringSelection in i['19:20 - 20:30']:
                        aggiunngimateria = funzioni.contalezioni(i['19:20 - 20:30'])
                        if aggiunngimateria == 1:
                            kokugotemp = kokugotemp +1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2, str(kokugotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 2, wx.GREEN)
                        if aggiunngimateria == 2:
                            eigotemp = eigotemp+1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0, str(eigotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 0, wx.GREEN)
                        if aggiunngimateria == 3:
                            suugakutemp = suugakutemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1, str(suugakutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 1, wx.GREEN)
                        if aggiunngimateria == 4:
                            rikatemp = rikatemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3, str(rikatemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 3, wx.GREEN)
                        if aggiunngimateria == 5:
                            shakaitemp = shakaitemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4, str(shakaitemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 4, wx.GREEN)
                        if aggiunngimateria == 6:
                            tokubetsutemp = tokubetsutemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5, str(tokubetsutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 5, wx.GREEN)

                    if self.listaStudenti.StringSelection in i['20:40 - 21:50']:
                        aggiunngimateria = funzioni.contalezioni(i['20:40 - 21:50'])
                        if aggiunngimateria == 1:
                            kokugotemp = kokugotemp +1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 2, str(kokugotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 2, wx.GREEN)
                        if aggiunngimateria == 2:
                            eigotemp = eigotemp+1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 0, str(eigotemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 0, wx.GREEN)
                        if aggiunngimateria == 3:
                            suugakutemp = suugakutemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 1, str(suugakutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 1, wx.GREEN)
                        if aggiunngimateria == 4:
                            rikatemp = rikatemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 3, str(rikatemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 3, wx.GREEN)
                        if aggiunngimateria == 5:
                            shakaitemp = shakaitemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 4, str(shakaitemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 4, wx.GREEN)
                        if aggiunngimateria == 6:
                            tokubetsutemp = tokubetsutemp + 1
                            self.grigliaLezioniSingole.SetCellValue(int(files[3:5]) - 1, 5, str(tokubetsutemp))
                            self.grigliaLezioniSingole.SetCellBackgroundColour(int(files[3:5]) - 1, 5, wx.GREEN)

                if u'K '+self.listaStudenti.StringSelection in i['9:10 - 10:20']:
                    contaPrivate= contaPrivate+1
                if u'K '+self.listaStudenti.StringSelection in i['10:30 - 11:40']:
                    contaPrivate= contaPrivate+1
                if u'K '+self.listaStudenti.StringSelection in i['11:50 - 13:00']:
                    contaPrivate= contaPrivate+1
                if u'K '+self.listaStudenti.StringSelection in i['13:40 - 14:50']:
                    contaPrivate= contaPrivate+1
                if u'K '+self.listaStudenti.StringSelection in i['15:00 - 16:10']:
                    contaPrivate= contaPrivate+1
                if u'K '+self.listaStudenti.StringSelection in i['16:40 - 17:50']:
                    contaPrivate= contaPrivate+1
                if u'K '+self.listaStudenti.StringSelection in i['18:00 - 19:10']:
                    contaPrivate= contaPrivate+1
                if u'K '+self.listaStudenti.StringSelection in i['19:20 - 20:30']:
                    contaPrivate= contaPrivate+1
                if u'K '+self.listaStudenti.StringSelection in i['20:40 - 21:50']:
                    contaPrivate= contaPrivate+1
                #print contaPrivate , 'contaprivate'
                sommainglese = 0
                sommakokugo = 0
                sommashakai = 0
                sommarika = 0
                sommatokubetsu = 0
                sommasuugaku = 0
                sommainglese1 = 0
                sommakokugo1 = 0
                sommashakai1 = 0
                sommarika1 = 0
                sommatokubetsu1 = 0
                sommasuugaku1 = 0
                for i in range(0, 31, 1):
                    if self.grigliaLezioniSingole.GetCellValue(i, 0) != u'' and self.grigliaLezioniSingole.GetCellBackgroundColour(i,0)==wx.GREEN:
                        #print self.grigliaLezioniSingole.GetCellBackgroundColour(i,0)
                        conv = self.grigliaLezioniSingole.GetCellValue(i, 0)
                        sommainglese = sommainglese + int(conv)
                for i in range(0, 31, 1):
                    if self.grigliaLezioniSingole.GetCellValue(i, 1) != u''and self.grigliaLezioniSingole.GetCellBackgroundColour(i,1)==wx.GREEN:
                        conv = self.grigliaLezioniSingole.GetCellValue(i, 1)
                        sommasuugaku = sommasuugaku + int(conv)
                for i in range(0, 31, 1):
                    if self.grigliaLezioniSingole.GetCellValue(i, 2) != u''and self.grigliaLezioniSingole.GetCellBackgroundColour(i,2)==wx.GREEN:
                        conv = self.grigliaLezioniSingole.GetCellValue(i, 2)
                        sommakokugo = sommakokugo + int(conv)
                for i in range(0, 31, 1):
                    if self.grigliaLezioniSingole.GetCellValue(i, 3) != u''and self.grigliaLezioniSingole.GetCellBackgroundColour(i,3)==wx.GREEN:
                        conv = self.grigliaLezioniSingole.GetCellValue(i, 3)
                        sommarika = sommarika + int(conv)
                for i in range(0, 31, 1):
                    if self.grigliaLezioniSingole.GetCellValue(i, 4) != u''and self.grigliaLezioniSingole.GetCellBackgroundColour(i,4)==wx.GREEN:
                        conv = self.grigliaLezioniSingole.GetCellValue(i, 4)
                        sommashakai = sommashakai + int(conv)
                for i in range(0, 31, 1):
                    if self.grigliaLezioniSingole.GetCellValue(i, 5) != u''and self.grigliaLezioniSingole.GetCellBackgroundColour(i,5)==wx.GREEN:
                        conv = self.grigliaLezioniSingole.GetCellValue(i, 5)
                        sommatokubetsu = sommatokubetsu + int(conv)

                for i in range(0, 31, 1):
                    if self.grigliaLezioniSingole.GetCellValue(i, 0) != u'' and self.grigliaLezioniSingole.GetCellBackgroundColour(i,0)==wx.WHITE:
                        #print self.grigliaLezioniSingole.GetCellBackgroundColour(i,0)
                        conv = self.grigliaLezioniSingole.GetCellValue(i, 0)
                        sommainglese1 = sommainglese1 + int(conv)
                for i in range(0, 31, 1):
                    if self.grigliaLezioniSingole.GetCellValue(i, 1) != u''and self.grigliaLezioniSingole.GetCellBackgroundColour(i,1)==wx.WHITE:
                        conv = self.grigliaLezioniSingole.GetCellValue(i, 1)
                        sommasuugaku1 = sommasuugaku1 + int(conv)
                for i in range(0, 31, 1):
                    if self.grigliaLezioniSingole.GetCellValue(i, 2) != u''and self.grigliaLezioniSingole.GetCellBackgroundColour(i,2)==wx.WHITE:
                        conv = self.grigliaLezioniSingole.GetCellValue(i, 2)
                        sommakokugo1 = sommakokugo1 + int(conv)
                for i in range(0, 31, 1):
                    if self.grigliaLezioniSingole.GetCellValue(i, 3) != u''and self.grigliaLezioniSingole.GetCellBackgroundColour(i,3)==wx.WHITE:
                        conv = self.grigliaLezioniSingole.GetCellValue(i, 3)
                        sommarika1 = sommarika1 + int(conv)
                for i in range(0, 31, 1):
                    if self.grigliaLezioniSingole.GetCellValue(i, 4) != u''and self.grigliaLezioniSingole.GetCellBackgroundColour(i,4)==wx.WHITE:
                        conv = self.grigliaLezioniSingole.GetCellValue(i, 4)
                        sommashakai1 = sommashakai1 + int(conv)
                for i in range(0, 31, 1):
                    if self.grigliaLezioniSingole.GetCellValue(i, 5) != u''and self.grigliaLezioniSingole.GetCellBackgroundColour(i,5)==wx.WHITE:
                        conv = self.grigliaLezioniSingole.GetCellValue(i, 5)
                        sommatokubetsu1 = sommatokubetsu1 + int(conv)
                sommarika1 = sommarika1 + int(sommarika)
                sommakokugo1 = sommakokugo1 + int(sommakokugo)
                sommasuugaku1 = sommasuugaku1 + int(sommasuugaku)
                sommainglese1 = sommainglese1 + int(sommainglese)
                sommashakai1 = sommashakai1 + int(sommashakai)

                sommatokubetsu1 = sommatokubetsu1 + int(sommatokubetsu)

                self.grigliaLezioniSingole.SetCellValue(32,0,str(contaPrivate))
                self.grigliaLezioniSingole.SetCellValue(31, 0, str(sommainglese))
                self.grigliaLezioniSingole.SetCellValue(31, 1, str(sommasuugaku))
                self.grigliaLezioniSingole.SetCellValue(31, 2, str(sommakokugo))
                self.grigliaLezioniSingole.SetCellValue(31, 3, str(sommarika))
                self.grigliaLezioniSingole.SetCellValue(31, 4, str(sommashakai))
                self.grigliaLezioniSingole.SetCellValue(31, 5, str(sommatokubetsu))
                self.grigliaLezioniSingole.SetCellValue(34, 0, str(sommainglese1))
                self.grigliaLezioniSingole.SetCellValue(34, 1, str(sommasuugaku1))
                self.grigliaLezioniSingole.SetCellValue(34, 2, str(sommakokugo1))
                self.grigliaLezioniSingole.SetCellValue(34, 3, str(sommarika1))
                self.grigliaLezioniSingole.SetCellValue(34, 4,  str(sommashakai1))
                self.grigliaLezioniSingole.SetCellValue(34, 5, str(sommatokubetsu1))
                balanceEigo = int(self.grigliaLezioniSingole.GetCellValue(34, 0)) - sommainglese
                balancesuugaku = int(self.grigliaLezioniSingole.GetCellValue(34, 1)) - sommasuugaku
                balancekokugo = int(self.grigliaLezioniSingole.GetCellValue(34, 2))- sommakokugo
                balancerika = int(self.grigliaLezioniSingole.GetCellValue(34, 3))  - sommarika
                balanceshakai = int(self.grigliaLezioniSingole.GetCellValue(34, 4))  - sommashakai
                balancetokubetu = int(self.grigliaLezioniSingole.GetCellValue(34, 5))  - sommatokubetsu
                self.grigliaLezioniSingole.SetCellValue(33, 0, str(balanceEigo))
                self.grigliaLezioniSingole.SetCellValue(33, 1, str(balancesuugaku))
                self.grigliaLezioniSingole.SetCellValue(33, 2, str(balancekokugo))
                self.grigliaLezioniSingole.SetCellValue(33, 3, str(balancerika))
                self.grigliaLezioniSingole.SetCellValue(33, 4, str(balanceshakai))
                self.grigliaLezioniSingole.SetCellValue(33, 5, str(balancetokubetu))
                if self.grigliaTotali.GetCellValue(0, 0) == '':
                    self.grigliaTotali.SetCellValue(0, 0,'0')
                if self.grigliaTotali.GetCellValue(0, 1) == '':
                    self.grigliaTotali.SetCellValue(0, 1,'0')
                if self.grigliaTotali.GetCellValue(0, 2) == '':
                    self.grigliaTotali.SetCellValue(0, 2,'0')
                if self.grigliaTotali.GetCellValue(0, 3) == '':
                    self.grigliaTotali.SetCellValue(0, 3,'0')
                if self.grigliaTotali.GetCellValue(0, 4) == '':
                    self.grigliaTotali.SetCellValue(0, 4,'0')
                if self.grigliaTotali.GetCellValue(0, 5) == '':
                    self.grigliaTotali.SetCellValue(0, 5,'0')
                tsuikaeigo = int(self.grigliaLezioniSingole.GetCellValue(34,0))-int(self.grigliaTotali.GetCellValue(0, 0))-int(self.grigliaTotali.GetCellValue(1, 0))
                tsuikakokugo = int(self.grigliaLezioniSingole.GetCellValue(34,2) ) - int(self.grigliaTotali.GetCellValue(0, 2)) - int(self.grigliaTotali.GetCellValue(1, 2))
                tsuikasuugaku = int(self.grigliaLezioniSingole.GetCellValue(34, 1)) - int(self.grigliaTotali.GetCellValue(0,1)) - int(self.grigliaTotali.GetCellValue(1, 1))
                tsuikarika = int(self.grigliaLezioniSingole.GetCellValue(34, 3)) - int(self.grigliaTotali.GetCellValue(0, 3)) - int(self.grigliaTotali.GetCellValue(1, 3))
                tsuikashakai = int(self.grigliaLezioniSingole.GetCellValue(34, 4)) - int(self.grigliaTotali.GetCellValue(0, 4)) - int(self.grigliaTotali.GetCellValue(1, 4))
                tsuikatokubetsu = int(self.grigliaLezioniSingole.GetCellValue(34, 5)) - int(self.grigliaTotali.GetCellValue(0, 5)) - int(self.grigliaTotali.GetCellValue(1, 5))
                if tsuikaeigo >= 0:
                    self.grigliaTotali.SetCellValue(2, 0, str(tsuikaeigo))
                else:
                    self.grigliaTotali.SetCellValue(2, 0, '0')
                if tsuikasuugaku >= 0:
                    self.grigliaTotali.SetCellValue(2, 1, str(tsuikasuugaku))
                else:
                    self.grigliaTotali.SetCellValue(2, 1, '0')
                if tsuikakokugo >= 0:
                    self.grigliaTotali.SetCellValue(2, 2, str(tsuikakokugo))
                else:
                    self.grigliaTotali.SetCellValue(2, 2, '0')
                if tsuikarika >= 0:
                    self.grigliaTotali.SetCellValue(2, 3, str(tsuikarika))
                else:
                    self.grigliaTotali.SetCellValue(2, 3, '0')
                if tsuikashakai >= 0:
                    self.grigliaTotali.SetCellValue(2, 4, str(tsuikashakai))
                else:
                    self.grigliaTotali.SetCellValue(2, 4, '0')
                if tsuikatokubetsu >= 0:
                    self.grigliaTotali.SetCellValue(2, 5, str(tsuikatokubetsu))
                else:
                    self.grigliaTotali.SetCellValue(2, 5, '0')
    def aggiornaDati(self, event):
        orario = {}
        orario1 = {}
        for creaorariofasullo in range(0, 9, 1):
            orario[creaorariofasullo] = False
            #print orario[creaorariofasullo]
        for creaorariofasullo in range(0, 9, 1):
            orario1[creaorariofasullo] = False
            #print orario[creaorariofasullo]
        for i in self.tabellaOre.Selections:
            #print len(self.tabellaOre.Items)
            orario[i] = True
            #print 'orarioi', orario[i]
        for i in self.tabellaOre1.Selections:
            #print len(self.tabellaOre.Items)
            orario1[i] = True
            #print orario[i]

        dati = dict(id=idSelezionato, name=self.casellaNome.Value, cardID=self.cardid.Label,
                    telephone=self.casellaTelefono.Value,
                    email=self.casellaEmail.Value, student=1,scuola=self.casellaScuola.Value,
                     maschio=self.danseiBox.Value, sonota=self.sonota.Value, femmina=self.joseiBox.Value,
                    furigana=self.furigana.Value,
                    parentMail=self.mailGenitori.Value,
                    teacher=0, kokugo=self.kokugo.Value, eigo=self.eigo.Value, suugaku=self.suugaku.Value,
                    rika=self.rika.Value, shakai=self.shakai.Value, tokubetsu=self.tokubetsu.Value,
                    primaOra=orario[0], secondaOra=orario[1], terzaOra=orario[2], quartaOra=orario[3],
                    quintaOra=orario[4], sestaOra=orario[5], settimaOra=orario[6], ottavaOra=orario[7],
                    nonaOra=orario[8],individual=self.individual.Value, shared=self.shared.Value,
                    kokugo1=self.kokugo1.Value, eigo1=self.eigo1.Value, suugaku1=self.suugaku1.Value,
                     rika1=self.rika1.Value, shakai1=self.shakai1.Value, tokubetsu1=self.tokubetsu1.Value,
                     primaOra1=orario1[0], secondaOra1=orario1[1], terzaOra1=orario1[2], quartaOra1=orario1[3],
                     quintaOra1=orario1[4], sestaOra1=orario1[5], settimaOra1=orario1[6], ottavaOra1=orario1[7],
                     nonaOra1=orario1[8])
        tabella.update(dati, ['id'])
        datigiorni = (dict(id = idGiorni, name=self.casellaNome.Value, lunedi=self.lunedi.Value,
                                  martedi=self.martedi.Value, mercoledi=self.mercoledi.Value,
                                  giovedi=self.giovedi.Value, venerdi=self.venerdi.Value,
                                  sabato=self.sabato.Value, domenica=self.domenica.Value,lunedi1=self.lunedi1.Value,
                                     martedi1=self.martedi1.Value, mercoledi1=self.mercoledi1.Value,
                                     giovedi1=self.giovedi1.Value, venerdi1=self.venerdi1.Value,
                                     sabato1=self.sabato1.Value, domenica1=self.domenica1.Value))

        tabellaGiorni.update(datigiorni,['id'])
        popolaCalcoli = tabellaCalcoli.find_one(name=self.listaStudenti.StringSelection,
                                                anno=self.calendarioStudenti.Date.Year,
                                                mese=self.calendarioStudenti.Date.Month)
        if popolaCalcoli is not None:
            datiCalcoli= (dict(id = popolaCalcoli.id,name=self.casellaNome.Value, mese=self.calendarioStudenti.Date.Month, anno=self.calendarioStudenti.Date.Year, normaleigo=self.grigliaTotali.GetCellValue(0,0),
                                           normalsuugaku=self.grigliaTotali.GetCellValue(0,1),normalkokugo=self.grigliaTotali.GetCellValue(0,2),
                                           normalrika=self.grigliaTotali.GetCellValue(0, 3),normalshakai=self.grigliaTotali.GetCellValue(0,4),
                                           normaltokubetsu=self.grigliaTotali.GetCellValue(0,5),
                                           tsuikaeigo=self.grigliaTotali.GetCellValue(2, 0),
                                           tsuikasuugaku=self.grigliaTotali.GetCellValue(2, 1),
                                           tsuikakokugo=self.grigliaTotali.GetCellValue(2, 2),
                                           tsuikarika=self.grigliaTotali.GetCellValue(2, 3),
                                           tsuikashakai=self.grigliaTotali.GetCellValue(2, 4),
                                           tsuikatokubetsu=self.grigliaTotali.GetCellValue(2, 5),
                                            # balanceeigo = self.grigliaLezioniSingole.GetCellValue(33, 0),
                                            # balancesuugaku = self.grigliaLezioniSingole.GetCellValue(33, 1),
                                            # balancekokugo = self.grigliaLezioniSingole.GetCellValue(33, 2),
                                            # balancerika = self.grigliaLezioniSingole.GetCellValue(33, 3),
                                            # balanceshakai = self.grigliaLezioniSingole.GetCellValue(33, 4),
                                            # balancetokubetu = self.grigliaLezioniSingole.GetCellValue(33, 5)
                               ))




            tabellaCalcoli.update(datiCalcoli,['id'])
        if popolaCalcoli is None:
            tabellaCalcoli.insert(dict(name=self.casellaNome.Value, anno=self.calendarioStudenti.Date.Year,
                                       mese=self.calendarioStudenti.Date.Month,
                                       normaleigo=self.grigliaTotali.GetCellValue(0, 0),
                                       normalsuugaku=self.grigliaTotali.GetCellValue(0, 1),
                                       normalkokugo=self.grigliaTotali.GetCellValue(0, 2),
                                       normalrika=self.grigliaTotali.GetCellValue(0, 3),
                                       normalshakai=self.grigliaTotali.GetCellValue(0, 4),
                                       normaltokubetsu=self.grigliaTotali.GetCellValue(0, 5),
                                       tsuikaeigo=self.grigliaTotali.GetCellValue(2, 0),
                                       tsuikasuugaku=self.grigliaTotali.GetCellValue(2, 1),
                                       tsuikakokugo=self.grigliaTotali.GetCellValue(2, 2),
                                       tsuikarika=self.grigliaTotali.GetCellValue(2, 3),
                                       tsuikashakai=self.grigliaTotali.GetCellValue(2, 4),
                                       tsuikatokubetsu=self.grigliaTotali.GetCellValue(2, 5),
                                       # balanceeigo = self.grigliaLezioniSingole.GetCellValue(33, 0),
                                       # balancesuugaku = self.grigliaLezioniSingole.GetCellValue(33, 1),
                                       # balancekokugo = self.grigliaLezioniSingole.GetCellValue(33, 2),
                                       # balancerika = self.grigliaLezioniSingole.GetCellValue(33, 3),
                                       # balanceshakai = self.grigliaLezioniSingole.GetCellValue(33, 4),
                                       # balancetokubetu = self.grigliaLezioniSingole.GetCellValue(33, 5)
                                       ))
        nomefile = './StudentsData/'+self.casellaNome.Value+self.casellaTelefono.Value+'.txt'
        nomefile1 = './StudentsData/' + self.casellaNome.Value + self.casellaTelefono.Value+'tokubetsu.txt'
        with open(nomefile, 'wb') as f:
            fieldnames = ['月曜日', '火曜日', '水曜日', '木曜日', '金曜日','土曜日', '日曜日']
            writer = csv.DictWriter(f, fieldnames=fieldnames, dialect='excel')

            writer.writeheader()
            for i in range(0,9 , 1):
                #print i

                #ciao =  utf_8_encoder(self.O.GetCellValue(i, 0))
                #print ciao, 'ciao'
                writer.writerow(
                    {'月曜日': utf_8_encoder(self.oreMaterie.GetCellValue(i, 0)), '火曜日':  utf_8_encoder(self.oreMaterie.GetCellValue(i, 1))
                        , '水曜日':  utf_8_encoder(self.oreMaterie.GetCellValue(i, 2)),
                     '木曜日':  utf_8_encoder(self.oreMaterie.GetCellValue(i, 3))
                        , '金曜日':  utf_8_encoder(self.oreMaterie.GetCellValue(i, 4)),
                     '土曜日':  utf_8_encoder(self.oreMaterie.GetCellValue(i, 5))
                        , '日曜日':  utf_8_encoder(self.oreMaterie.GetCellValue(i, 6))})
        with open(nomefile1, 'wb') as f:
            fieldnames = ['月曜日', '火曜日', '水曜日', '木曜日', '金曜日','土曜日', '日曜日']
            writer = csv.DictWriter(f, fieldnames=fieldnames, dialect='excel')

            writer.writeheader()
            for i in range(0,9 , 1):
                #print i

                #ciao =  utf_8_encoder(self.O.GetCellValue(i, 0))
                #print ciao, 'ciao'
                writer.writerow(
                    {'月曜日': utf_8_encoder(self.oreMaterie1.GetCellValue(i, 0)), '火曜日':  utf_8_encoder(self.oreMaterie1.GetCellValue(i, 1))
                        , '水曜日':  utf_8_encoder(self.oreMaterie1.GetCellValue(i, 2)),
                     '木曜日':  utf_8_encoder(self.oreMaterie1.GetCellValue(i, 3))
                        , '金曜日':  utf_8_encoder(self.oreMaterie1.GetCellValue(i, 4)),
                     '土曜日':  utf_8_encoder(self.oreMaterie1.GetCellValue(i, 5))
                        , '日曜日':  utf_8_encoder(self.oreMaterie1.GetCellValue(i, 6))})

        for i in self.usciteStudenti.Items:
            if tabellaTempo.find_one(name=self.casellaNome.Value, uscite=i)is None:
                tabellaTempo.insert(dict(name=self.casellaNome.Value, uscite=i))
        self.listaStudenti.Clear()
        popolaStudenti = tabella.find(student='1')
        for i in popolaStudenti:
            self.listaStudenti.Append(unicode(i.name))
    def abilitaNuovo(self, event):
        self.invio.Enabled = True
        self.cardid.LabelText = ''
    def cancellaDati( self, event ):
        dlg = wx.MessageDialog(None, u"データ削除しますか", '', wx.YES_NO | wx.ICON_QUESTION)
        result = dlg.ShowModal()

        if result == wx.ID_YES:

            tabella.delete(name=self.casellaNome.Value)
            tabellaTempo.delete(name=self.casellaNome.Value)
            tabellaGiorni.delete(name=self.casellaNome.Value)
            tabellaCalcoli.delete(name=self.casellaNome.Value)
            self.listaStudenti.Clear()
            popolaStudenti = tabella.find(student='1')
            for i in popolaStudenti:
                self.listaStudenti.Append(unicode(i.name))
        else:
            pass
            #print "No pressed"
    def aggiungiData(self, event):
        calendario = calendar

        giornoDelMeseCorrente = str(self.calendarioStudenti.Date)
        dataDatetime = datetime.strptime(giornoDelMeseCorrente, '%m/%d/%y %H:%M:%S')
        lungezzaMese = calendario.monthrange(dataDatetime.year,dataDatetime.month)

        dataComp = str(self.calendarioStudenti.Date.Year)+'/'+str(self.calendarioStudenti.Date.Month+1)+'/'+str(self.calendarioStudenti.Date.Day)
        dataComposta = funzioni.aggiungizeri(self.calendarioStudenti.Date.Year,self.calendarioStudenti.Date.Month+1,self.calendarioStudenti.Date.Day)
        print dataComposta
        controllaDate = funzioni.controlloDateDuplicate(dataComposta, self.usciteStudenti.Items)
        print type (self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(self.calendarioStudenti.Date)
        if controllaDate == True:
            self.usciteStudenti.Append(dataComposta)
        else:
            self.errorCheck.LabelText = u'すでに追加されました'
    def cancellaLezioni(self, event):
        # tabellaTempo.delete(name=self.casellaNome.LabelText,riposi=self.riposiInsegnanti.GetSelections)
        selections = list(self.usciteStudenti.GetSelections())
        datadaCancellare = self.usciteStudenti.StringSelection
        for index in selections:

            #print self.casellaNome.LabelText, self.usciteStudenti.StringSelection, type(self.listaStudenti.LabelText), type (self.riposiInsegnanti.StringSelection)
            self.usciteStudenti.Delete(index)
        tabellaTempo.delete(name=self.listaStudenti.StringSelection,uscite=datadaCancellare)
    def readCard( self, event ):
        idcarta = cardScan()
        cercacarta = tabella.find_one(cardID=idcarta)
        if cercacarta is not None:
            self.errorCheck.LabelText = 'Card Already on database'
        else:
            self.errorCheck.LabelText = 'Card reading successfully'
            self.cardid.Label = idcarta
    def aggiuntaGiorni( self, event ):


        giornoDelMeseCorrente = str(self.calendarioStudenti.Date)
        dataDatetime = datetime.strptime(giornoDelMeseCorrente, '%m/%d/%y %H:%M:%S')
        print dataDatetime
        calendario = calendar
        print calendario.month(dataDatetime.year,dataDatetime.month)
        print calendario.monthrange(dataDatetime.year,dataDatetime.month), type (calendario.monthrange(dataDatetime.year,dataDatetime.month))
        lungezzaMese = calendario.monthrange(dataDatetime.year,dataDatetime.month)
        if self.lunedi.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,0,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.usciteStudenti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.usciteStudenti.Append(i)
        if self.martedi.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,1,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.usciteStudenti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.usciteStudenti.Append(i)
        if self.mercoledi.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,2,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.usciteStudenti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.usciteStudenti.Append(i)
        if self.giovedi.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,3,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.usciteStudenti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.usciteStudenti.Append(i)
        if self.venerdi.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,4,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.usciteStudenti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.usciteStudenti.Append(i)
        if self.sabato.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,5,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.usciteStudenti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.usciteStudenti.Append(i)
        if self.domenica.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,6,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.usciteStudenti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.usciteStudenti.Append(i)
    def aggiuntaGiorni1( self, event ):


        giornoDelMeseCorrente = str(self.calendarioStudenti.Date)
        dataDatetime = datetime.strptime(giornoDelMeseCorrente, '%m/%d/%y %H:%M:%S')
        print dataDatetime
        calendario = calendar
        print calendario.month(dataDatetime.year,dataDatetime.month)
        print calendario.monthrange(dataDatetime.year,dataDatetime.month), type (calendario.monthrange(dataDatetime.year,dataDatetime.month))
        lungezzaMese = calendario.monthrange(dataDatetime.year,dataDatetime.month)
        if self.lunedi1.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,0,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.usciteStudenti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.usciteStudenti.Append(i)
        if self.martedi1.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,1,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.usciteStudenti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.usciteStudenti.Append(i)
        if self.mercoledi1.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,2,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.usciteStudenti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.usciteStudenti.Append(i)
        if self.giovedi1.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,3,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.usciteStudenti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.usciteStudenti.Append(i)
        if self.venerdi1.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,4,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.usciteStudenti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.usciteStudenti.Append(i)
        if self.sabato1.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,5,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.usciteStudenti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.usciteStudenti.Append(i)
        if self.domenica1.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,6,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.usciteStudenti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.usciteStudenti.Append(i)


class finestraUtenti(JukuPlanner.addUser):
    # constructor
    def __init__(self, parent):
        # initialize parent class
        JukuPlanner.addUser.__init__(self, parent)
        popolaInsegnanti = tabellaIns.find(teacher='1')
        for i in popolaInsegnanti:
            self.listaInsegnanti.Append(unicode(i.name))
        self.invio.Enabled = True
        self.aggiorna.Enabled = False
        self.cancella.Enabled = False
    def mostraMeseCorrente( self, event ):
        listadate = []
        dataComp = str(self.calendarioStudenti.Date.Year) + '/' + str(self.calendarioStudenti.Date.Month + 1) + '/'
        dataComposta = funzioni.aggiungizeriSenzaGiorno(self.calendarioStudenti.Date.Year, self.calendarioStudenti.Date.Month + 1)
        dataunicode = unicode(dataComposta)
        contaitem = 0
        popolaDate = tabellaTempoIns.find(name=self.listaInsegnanti.StringSelection)

        if self.meseCorrente.Value == True:
            for i in self.riposiInsegnanti.Items:
                if dataunicode in i :
                    listadate.append(i)
            self.riposiInsegnanti.Clear()
            for i in listadate:
                self.riposiInsegnanti.Append(i)
        if self.meseCorrente.Value == False:
            self.riposiInsegnanti.Clear()
            for i in popolaDate:
                if len((str(i.riposi))) >= 5:
                    self.riposiInsegnanti.Append(str(i.riposi))

    def cancellaRiposi(self, event):
        # tabellaTempo.delete(name=self.casellaNome.LabelText,riposi=self.riposiInsegnanti.GetSelections)
        selections = list(self.riposiInsegnanti.GetSelections())

        for index in selections:

            print self.casellaNome.LabelText, self.riposiInsegnanti.StringSelection, type(self.casellaNome.LabelText), type (self.riposiInsegnanti.StringSelection)
            self.riposiInsegnanti.Delete(index)
        tabellaTempoIns.delete(name=self.casellaNome.LabelText)
    def nuovoInsegnante( self, event ):

        self.invio.Enabled = True
        self.aggiorna.Enabled = False
        self.cancella.Enabled = False
        self.casellaNome.Clear()
        self.casellaTelefono.Clear()
        self.casellaEmail.Clear()
        self.tabellaOre.DeselectAll()
        self.tabellaOre1.DeselectAll()
        self.riposiInsegnanti.Clear()
        self.furigana.Clear()
        self.lunedi.Value = False
        self.martedi.Value = False
        self.mercoledi.Value = False
        self.giovedi.Value = False
        self.venerdi.Value = False
        self.sabato.Value = False
        self.domenica.Value = False
        self.listaInsegnanti.Clear()
        popolaStudenti = tabellaIns.find(teacher='1')
        for i in popolaStudenti:
            self.listaInsegnanti.Append(unicode(i.name))

    def orePersonalizzate( self, event ):
        print str(self.calendarioStudenti.Date)
        popolaDateIns = tabellaDateIns.find_one(name=self.casellaNome.Value, data=unicode(self.calendarioStudenti.Date))
        popolainsegnanti = tabellaIns.find_one(name=self.listaInsegnanti.StringSelection)

        if popolaDateIns is not None:
            arrayore = [popolaDateIns.primaOra,popolaDateIns.secondaOra,popolaDateIns.terzaOra,popolaDateIns.quartaOra,popolaDateIns.quintaOra,popolaDateIns.sestaOra,popolaDateIns.settimaOra,popolaDateIns.ottavaOra,popolaDateIns.nonaOra]
            for settaOre in range(0, 9, 1):
                self.tabellaOre1.SetSelection(settaOre, select=arrayore[settaOre])
        if popolaDateIns is  None:

            arrayore = [popolainsegnanti.primaOra, popolainsegnanti.secondaOra, popolainsegnanti.terzaOra, popolainsegnanti.quartaOra, popolainsegnanti.quintaOra, popolainsegnanti.sestaOra, popolainsegnanti.settimaOra,
                        popolainsegnanti.ottavaOra, popolainsegnanti.nonaOra]
            for settaOre in range(0, 9, 1):
                self.tabellaOre1.SetSelection(settaOre, select=arrayore[settaOre])
    def caricaDate(self, event):
        self.errorCheck.LabelText='-------------------------------------------------------------------------------------------------------------------------------------------------'
        self.riposiInsegnanti.Clear()
        self.aggiorna.Enabled = True
        self.cancella.Enabled = True
        self.invio.Enabled = False
        print self.listaInsegnanti.StringSelection
        popolaDateIns = tabellaDateIns.find_one(name=self.casellaNome.Value, data=str(self.calendarioStudenti.Date))
        popolaDate = tabellaTempoIns.find(name=self.listaInsegnanti.StringSelection)
        popolainsegnanti = tabellaIns.find(name=self.listaInsegnanti.StringSelection, teacher='1')
        global idSelezionato
        global  idDatePersonalizzate
        if popolaDateIns is not None:
            idDatePersonalizzate = popolaDateIns.id
        for i in popolainsegnanti:

            idSelezionato = i.id
            self.lunedi.Value = i.lunedi
            self.martedi.Value = i.martedi
            self.mercoledi.Value = i.mercoledi
            self.giovedi.Value = i.giovedi
            self.venerdi.Value = i.venerdi
            self.sabato.Value = i.sabato
            self.domenica.Value = i.domenica
            self.casellaNome.LabelText = i.name
            self.casellaEmail.LabelText = i.email
            self.furigana.Value = i.furigana


            self.casellaTelefono.LabelText = i.telephone
            # self.eigo.Value = i.eigo
            # self.rika.Value = i.rika
            # self.shakai.Value = i.shakai
            # self.suugaku.Value = i.suugaku
            # self.tokubetsu.Value = i.tokubetsu
            # self.kokugo.Value = i.kokugo
            if i.cardID == '':
                self.cardid.LabelText = u"カード未登録です、登録してください"
                self.CardRegistration.Enabled=True

                self.cardcancel.Enabled = False
            else:
                self.cardid.LabelText = i.cardID
                self.CardRegistration.Enabled=False
                self.cardcancel.Enabled=True
            arrayore = [i.primaOra, i.secondaOra, i.terzaOra, i.quartaOra, i.quintaOra, i.sestaOra, i.settimaOra,
                        i.ottavaOra, i.nonaOra]
            print arrayore
            for settaOre in range(0, 9, 1):
                self.tabellaOre.SetSelection(settaOre, select=arrayore[settaOre])
            for settaOre in range(0, 9, 1):
                self.tabellaOre1.SetSelection(settaOre, select=arrayore[settaOre])

        for i in popolaDate:
            self.riposiInsegnanti.Append(str(i.riposi))

        self.invio.Enabled = False

    def cancellaDati(self, event):
        dlg = wx.MessageDialog(None, u"データ削除しますか", '', wx.YES_NO | wx.ICON_QUESTION)
        result = dlg.ShowModal()

        if result == wx.ID_YES:
            tabellaIns = dbins['insegnanti']
            tabellaTempoIns = dbins['timeTable']
            tabellaDateIns = dbins['datePersonalizzate']

            tabellaIns.delete(name=self.casellaNome.Value)
            tabellaTempoIns.delete(name=self.casellaNome.Value)
            tabellaDateIns.delete(name=self.casellaNome.Value)

            self.listaInsegnanti.Clear()
            popolaStudenti = tabellaIns.find(teacher='1')
            for i in popolaStudenti:
                self.listaInsegnanti.Append(unicode(i.name))
        else:

            pass
    def aggiornaDati(self, event):
        global idSelezionato
        global idDatePersonalizzate
        popolaDateIns = tabellaDateIns.find_one(name=self.casellaNome.Value, data=str(self.calendarioStudenti.Date))
        if popolaDateIns is not None:
            idDatePersonalizzate = popolaDateIns.id

        #idDatePersonalizzate = popolaDateIns.id
        orario = {}
        for creaorariofasullo in range(0, 9, 1):
            orario[creaorariofasullo] = False
        orario1 = {}
        for creaorariofasullo in range(0, 9, 1):
            orario1[creaorariofasullo] = False
        #print orario[creaorariofasullo]
        for i in self.tabellaOre.Selections:
            print len(self.tabellaOre.Items)
            orario[i] = True
            print 'orarioi', orario[i]
        for i in self.tabellaOre1.Selections:
            print len(self.tabellaOre1.Items)
            orario1[i] = True
        print orario, orario1, 'orari'
        if orario == orario1:
            dati = dict(id=idSelezionato, name=self.casellaNome.Value, cardID=self.cardid.Label,
                        telephone=self.casellaTelefono.Value,
                        email=self.casellaEmail.Value, student=0,
                        teacher=1,furigana = self.furigana.Value,
                        primaOra=orario[0], secondaOra=orario[1], terzaOra=orario[2], quartaOra=orario[3],
                        quintaOra=orario[4], sestaOra=orario[5], settimaOra=orario[6], ottavaOra=orario[7],
                        nonaOra=orario[8], lunedi=self.lunedi.Value,
                        martedi=self.martedi.Value, mercoledi=self.mercoledi.Value,
                        giovedi=self.giovedi.Value, venerdi=self.venerdi.Value,
                        sabato=self.sabato.Value, domenica=self.domenica.Value)
            tabellaIns.update(dati, ['id'])
        else:
            if popolaDateIns is None:
                print str(self.calendarioStudenti.Date)
                tabellaDateIns.insert(
                    dict(name=self.casellaNome.Value, data=str(self.calendarioStudenti.Date),
                         primaOra=orario1[0], secondaOra=orario1[1], terzaOra=orario1[2], quartaOra=orario1[3],
                         quintaOra=orario1[4], sestaOra=orario1[5], settimaOra=orario1[6], ottavaOra=orario1[7],
                         nonaOra=orario1[8]))
                dati = dict(id=idSelezionato, name=self.casellaNome.Value, cardID=self.cardid.Label,
                            telephone=self.casellaTelefono.Value,
                            email=self.casellaEmail.Value, student=0,
                            teacher=1,furigana = self.furigana.Value,
                            primaOra=orario[0], secondaOra=orario[1], terzaOra=orario[2], quartaOra=orario[3],
                            quintaOra=orario[4], sestaOra=orario[5], settimaOra=orario[6], ottavaOra=orario[7],
                            nonaOra=orario[8], lunedi=self.lunedi.Value,
                            martedi=self.martedi.Value, mercoledi=self.mercoledi.Value,
                            giovedi=self.giovedi.Value, venerdi=self.venerdi.Value,
                            sabato=self.sabato.Value, domenica=self.domenica.Value)
                tabellaIns.update(dati, ['id'])
            if popolaDateIns is not None:
                dati1 = dict(id=idDatePersonalizzate, name=self.casellaNome.Value,
                            primaOra=orario1[0], secondaOra=orario1[1], terzaOra=orario1[2], quartaOra=orario1[3],
                            quintaOra=orario1[4], sestaOra=orario1[5], settimaOra=orario1[6], ottavaOra=orario1[7],
                            nonaOra=orario1[8])
                tabellaDateIns.update(dati1, ['id'])
                dati = dict(id=idSelezionato, name=self.casellaNome.Value, cardID=self.cardid.Label,
                            telephone=self.casellaTelefono.Value,
                            email=self.casellaEmail.Value, student=0,
                            teacher=1,furigana = self.furigana.Value,
                            primaOra=orario[0], secondaOra=orario[1], terzaOra=orario[2], quartaOra=orario[3],
                            quintaOra=orario[4], sestaOra=orario[5], settimaOra=orario[6], ottavaOra=orario[7],
                            nonaOra=orario[8], lunedi=self.lunedi.Value,
                            martedi=self.martedi.Value, mercoledi=self.mercoledi.Value,
                            giovedi=self.giovedi.Value, venerdi=self.venerdi.Value,
                            sabato=self.sabato.Value, domenica=self.domenica.Value)
                tabellaIns.update(dati, ['id'])
            print 'diversi'

        for i in self.riposiInsegnanti.Items:
            if tabellaTempoIns.find_one(name=self.casellaNome.Value, riposi=i) is None:
                tabellaTempoIns.insert(dict(name=self.casellaNome.Value, riposi=i))
        #self.listaInsegnanti.Clear()
        #popolaStudenti = tabellaIns.find(teacher='1')
        #for i in popolaStudenti:

        #    self.listaInsegnanti.Append(unicode(i.name))


    def cardDelete(self, event):
        self.cardid.Label=''
        orario = {}
        for creaorariofasullo in range(0, 9, 1):
            orario[creaorariofasullo] = False
            print orario[creaorariofasullo]
        for i in self.tabellaOre.Selections:
            print len(self.tabellaOre.Items)
            orario[i] = True
            print 'orarioi', orario[i]

        dati = dict(id=idSelezionato, name=self.casellaNome.Value, cardID=self.cardid.Label,
                    telephone=self.casellaTelefono.Value,
                    email=self.casellaEmail.Value,furigana = self.furigana.Value,
                    teacher=1, kokugo=self.kokugo.Value, eigo=self.eigo.Value, suugaku=self.suugaku.Value,
                    rika=self.rika.Value, shakai=self.shakai.Value, tokubetsu=self.tokubetsu.Value,
                    primaOra=orario[0], secondaOra=orario[1], terzaOra=orario[2], quartaOra=orario[3],
                    quintaOra=orario[4], sestaOra=orario[5], settimaOra=orario[6], ottavaOra=orario[7],
                    nonaOra=orario[8])
        tabellaIns.update(dati, ['id'])
        for i in self.riposiInsegnanti.Items:
            if tabellaTempoIns.find_one(name=self.casellaNome.Value, riposi=i) is None:
                tabellaTempoIns.insert(dict(name=self.casellaNome.Value, riposi=i))

    def abilitaNuovo(self, event):
        self.invio.Enabled = True
        self.cardid.LabelText = ''
    def aggiuntaGiorni( self, event ):


        giornoDelMeseCorrente = str(self.calendarioStudenti.Date)
        dataDatetime = datetime.strptime(giornoDelMeseCorrente, '%m/%d/%y %H:%M:%S')
        print dataDatetime
        calendario = calendar
        print calendario.month(dataDatetime.year,dataDatetime.month)
        print calendario.monthrange(dataDatetime.year,dataDatetime.month), type (calendario.monthrange(dataDatetime.year,dataDatetime.month))
        lungezzaMese = calendario.monthrange(dataDatetime.year,dataDatetime.month)
        if self.lunedi.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,0,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.riposiInsegnanti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.riposiInsegnanti.Append(i)
        if self.martedi.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,1,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.riposiInsegnanti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.riposiInsegnanti.Append(i)
        if self.mercoledi.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,2,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.riposiInsegnanti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.riposiInsegnanti.Append(i)
        if self.giovedi.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,3,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.riposiInsegnanti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.riposiInsegnanti.Append(i)
        if self.venerdi.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,4,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.riposiInsegnanti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.riposiInsegnanti.Append(i)
        if self.sabato.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,5,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.riposiInsegnanti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.riposiInsegnanti.Append(i)
        if self.domenica.Value == True:
            tuplaGiorni = funzioni.giorniSettimana(dataDatetime.year,dataDatetime.month,6,lungezzaMese[1])
            for i in tuplaGiorni:
                controllaDate = funzioni.controlloDateDuplicate(i, self.riposiInsegnanti.Items)
                print type(self.calendarioStudenti.Date), type(str(self.calendarioStudenti.Date)), str(
                    self.calendarioStudenti.Date)
                if controllaDate == True:
                    self.riposiInsegnanti.Append(i)

    def aggiungiData(self, event):
        calendario = calendar

        giornoDelMeseCorrente = str(self.calendarioStudenti.Date)
        dataDatetime = datetime.strptime(giornoDelMeseCorrente, '%m/%d/%y %H:%M:%S')
        lungezzaMese = calendario.monthrange(dataDatetime.year, dataDatetime.month)
        dataComposta = funzioni.aggiungizeri(self.calendarioStudenti.Date.Year, self.calendarioStudenti.Date.Month + 1,
                                             self.calendarioStudenti.Date.Day)
        controllaDate = funzioni.controlloDateDuplicate(dataComposta, self.riposiInsegnanti.Items)
        if controllaDate == True:
            self.riposiInsegnanti.Append(dataComposta)
        else:
            self.errorCheck.LabelText = u'すでに追加されました'

    def selezioneCalendario(self, event):
        controllaDate = funzioni.controlloDateDuplicate(self.calendarioStudenti.Date, )
        self.text.SetValue(str(self.calendarioStudenti.Date))

    # put a blank string in text when 'Clear' is clicked
    def clearFunc(self, event):
        self.text.SetValue(str(''))

    def funzioneInvio(self, event):
        orario = {}
        for creaorariofasullo in range(0, 9, 1):
            orario[creaorariofasullo] = False
            print orario[creaorariofasullo]
        cercaNome = tabella.find_one(name=self.casellaNome.Value)
        print self.tabellaOre.Selections
        for i in self.tabellaOre.Selections:
            print len(self.tabellaOre.Items)
            orario[i] = True
            print orario[i]
        if cercaNome is not None:
            self.errorCheck.LabelText = 'Name Already on database'
        if self.casellaNome.Value is None:
            self.errorCheck.LabelText = 'Please fill the name'
        else:
            tabellaIns.insert(
                dict(name=self.casellaNome.Value, cardID=self.cardid.Label, telephone=self.casellaTelefono.Value,
                     email=self.casellaEmail.Value, student=0,furigana = self.furigana.Value,
                     teacher=1,lunedi=self.lunedi.Value,
                            martedi=self.martedi.Value, mercoledi=self.mercoledi.Value,
                            giovedi=self.giovedi.Value, venerdi=self.venerdi.Value,
                            sabato=self.sabato.Value, domenica=self.domenica.Value,
                     primaOra=orario[0], secondaOra=orario[1], terzaOra=orario[2], quartaOra=orario[3],
                     quintaOra=orario[4], sestaOra=orario[5], settimaOra=orario[6], ottavaOra=orario[7],
                     nonaOra=orario[8]))
            for i in self.riposiInsegnanti.Items:
                tabellaTempoIns.insert(dict(name=self.casellaNome.Value, riposi=i))
            print tabella
            self.errorCheck.LabelText = 'Data has been saved!'

    def noMaster(self, event):
        self.teacherCheckBox.Value = 0

    def noStudent(self, event):
        self.studentCheckBok.Value = 0

    def readCard(self, event):
        idcarta = cardScan()
        cercacarta = tabellaIns.find_one(cardID=idcarta)
        if cercacarta is not None:
            self.errorCheck.LabelText = 'Card Already on database'
        else:
            self.errorCheck.LabelText = 'Card reading successfully'
            self.cardid.Label = idcarta


def cardScan():
    cmd = ['python', 'tagtool.py']
    subprocess.Popen(cmd).wait()

    while True:
        quantiTxt = glob.glob("tag.txt")
        if len(quantiTxt) >= 1:
            filenfc = open('tag.txt', 'r')
            linea = filenfc.readlines()
            tagnfc = linea[0]
            # print tagnfc
            print ('Card reading complete')
            print tagnfc, 'Tagnfc'
            return tagnfc
            break
        else:
            time.sleep(1)


# mandatory in wx, create an app, False stands for not deteriction stdin/stdout
# refer manual for details
app = wx.App(False)

# create an object of CalcFrame
frame = CalcFrame(None)
# show the frame
frame.Show(True)
# start the applications
app.MainLoop()
