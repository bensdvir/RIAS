# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

# k fold  cross validation with avg
#
from docx import Document
import xlsxwriter
import numpy as np
import keras
import xlrd
import sys
from keras.models import Sequential
from keras.layers import Dense
from keras.layers import LSTM
import keras.preprocessing.text
from keras.preprocessing import sequence
from keras.layers.embeddings import Embedding
from sklearn.ensemble import RandomForestClassifier
from xlutils.copy import copy
from xlrd import open_workbook 
import random
from numpy import array
import networkx as nx
import matplotlib.pyplot as plt
from keras import backend as K




isRandom = False
tapeindexes=[None] * 14
doc = Document('C:\RIASmanual2016.docx')
try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile
WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
tags = ["Personal","Laughs", "Concern","R/O","Approve","Comp",
                "Disagree","Crit","Legit","Partner","Self-Dis","?Reassure",      
                "Agree","BC","Trans","Orient","Checks","?Understand","?Bid",
                "?Opinion","?Permission","? Understand","? Bid","Emp",
                "? Opinion","? Permission","Gives-Med","Gives-Thera","Gives-L/S",
                "Gives-P/S","Gives-Other","[?]Med","[?]Thera","[?]L/S","[?]P/S","[?]Other","[?] Med","[?] Thera","[?] L/S","[?] P/S","[?] Other ",
                "?Med","?Thera","?L/S","?P/S","?Other","C-Med/Thera","?Service","Unintel","C-L/S-P/S","? Med","? Thera","? L/S","? P/S","? Other"]


def get_docx_text(path):
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)

    paragraphs = []
    for paragraph in tree.getiterator(PARA):
        texts = [node.text
                 for node in paragraph.getiterator(TEXT)
                 if node.text]
        if texts:
            paragraphs.append(' '.join(texts))

    return '\n\n'.join(paragraphs)

def isTag(text):          
    cleaned = text.lstrip()
    cleaned = cleaned.rstrip()
    cleaned= cleaned.replace("- ","-")
    cleaned= cleaned.replace("[?]  ","[?] ")
    cleaned= cleaned.replace("[?] ","[?]")
    cleaned = cleaned.replace("?  ","? ")
    cleaned = cleaned.replace("? ","?")
    cleaned =  cleaned.split()
    for s in cleaned:
        if s not in tags:
            return False
    return True

#################################################### parsing ##############################################################################

def parseWordtoXExcel(path):
    row = 0
    col = 0            
    
    workbook =xlsxwriter.Workbook('C:\hi\data.xlsx')
    worksheet = workbook.add_worksheet()
    text = get_docx_text(path)
    splited = text.splitlines()    
    index = -1
    lineNum = 0
    tapenum =0
    n = 0
    lau = False
    for t in splited:
        index +=1
        if ("D." in t or "P." in t):
            lineNum+=1
    
            if ("/" in splited[index+2]):
                tempText = splited[index+2]
                tempTags = splited[index-2]
                if("(both" in tempText or "(lau" in tempText):
                    toCont = True
                    tempStr = tempText
                    halfInd = tempStr.index("(")
                    halfInd2 = tempStr.index(")")
                    if ("/ " in tempStr[halfInd-3:] or "/    " in tempStr[halfInd-15:]):
                         toCont = False
                    if (toCont):
                        if (0 == halfInd):
                            firstPart = ""
                        else:
                            firstPart = tempStr[:halfInd-1]
                        seconedPart = tempStr[halfInd:halfInd2+1]
                        thirdPart =tempStr[halfInd2+1:]
                        if (not ("" == firstPart)):
                            tempText= ""+str(firstPart)
                        if (not ("" == seconedPart)):
                            tempText+= "/ "+str(seconedPart)
                        if (not ("" == thirdPart)):
                            tempText+= "/ "+str(thirdPart)
                        lau = True
        
                    
                ind = 2
                while( index+2+ind < len(splited) and ("D." not in str(splited[index+2+ind]) and "P." not in str(splited[index+2+ind]))):
                    if(str(splited[index+2+ind]) is " "):
                        ind+=1
                        continue
                    if isTag(str(splited[index+2+ind])):
                        tempTags+= " " +str(splited[index+2+ind])
                    else:
                        if("(both" in splited[index+2+ind] or "(lau" in splited[index+2+ind] and not lau ):
                            tempStr = str(splited[index+2+ind])
                            halfInd = tempStr.index("(")
                            halfInd2 = tempStr.index(")")
                            if (not (halfInd == 0)):
                                if ("/ " in tempStr[halfInd-3:] or "/    " in tempStr[halfInd-15:]):
                                    ind+=2
                                    continue
                            if (0 == halfInd):
                                  firstPart = ""
                            else:
                                firstPart = tempStr[:halfInd-1]
                            seconedPart = tempStr[halfInd:halfInd2]
                            thirdPart =tempStr[halfInd2+1:]
                            if (not ("" == seconedPart)):
                                if((halfInd == 0)):
                                    tempText+= str(seconedPart)
                                else:
                                    tempText+= "/ "+str(seconedPart)
                            if (not ("" == thirdPart)):
                                tempText+= "/ "+str(thirdPart)
                            lau = True
                        else:
                            tempText+= " "+str(splited[index+2+ind])
                    ind+=2
                    lau = False 
                multi = tempText.split("/")
                for i in multi:
                    if(''==i or ' '==i or '  '==i):
                        multi.remove(i)
                tempTags= tempTags.replace("- ","-")
                tempTags= tempTags.replace("[?]  ","[?] ")
                tempTags= tempTags.replace("[?] ","[?]")
                tempTags = tempTags.replace("?  ","? ")
                tempTags = tempTags.replace("? ","?")
                tags =  tempTags.split()                        
                allText = ''
                i = 0
                for muli in multi:
                    allText+= "text"+str(i) + ": " + multi[i] +"  *****  "+ "tag" +str(i)+ ": " + tags[i] +"  *****  "
                    worksheet.write(row, col, lineNum)
                    if("P." in t):
                        worksheet.write(row, col + 1, "Patient")
                    else:
                        worksheet.write(row, col + 1, "Doctor")
                    tempSen = str(multi[i])
                    tempSen =tempSen.lstrip()
                    tempSen = tempSen.rstrip()
                    worksheet.write(row, col + 2, tempSen)
                    tg = str(tags[i])
                    tg= tg.replace("- ","-")
                    tg= tg.replace("[?]  ","[?] ")
                    tg= tg.replace("[?] ","[?]")
                    tg = tg.replace("?  ","? ")
                    tg = tg.replace("? ","?")
                    tg = tg.lstrip()
                    tg = tg.rstrip()
                    t1 = (str(tg)).replace(" ","")
                    t1 = t1.lower()
                    worksheet.write(row, col + 3,  t1)
                    i+=1
                    row+=1
                n+=len(multi)
            else: 
                tg = splited[index-2]
                tg= tg.replace("[?] ","[?]")
                tg = tg.replace("? ","?")
                tg = tg.lstrip()
                tg = tg.rstrip()
                allText = "Text: " + splited[index+2] + "  *****  " + "Tag: " + tg
                worksheet.write(row, col, lineNum)
                if("P." in t):
                    worksheet.write(row, col + 1, "Patient")
                else:
                    worksheet.write(row, col + 1, "Doctor") 
                tempSen = splited[index+2]
                tempSen =tempSen.lstrip()
                tempSen = tempSen.rstrip()
                if("(both" in tempSen or "(lau" in tempSen or "( lau" in tempSen):
                    lau = True
                    tempStr = tempSen
                    halfInd = tempStr.index("(")
                    halfInd2 = tempStr.index(")")               
                    halfIndTg = tg.index("Lau")
                    halfInd2Tg = halfIndTg+5
                    
                    if (0 == halfInd):
                        firstPart = ""
                    else:
                        firstPart = tempStr[:halfInd-1]
                    seconedPart = tempStr[halfInd:halfInd2+1]
                    thirdPart =tempStr[halfInd2+1:]
                    
                    firstPartTg = tg[:halfIndTg-1]
                    seconedPartTg= tg[halfIndTg:halfInd2Tg+1]
                    thirdPartTg = tg[halfInd2Tg+1:]
    
    
                    if (not ("" == firstPart)):
                        worksheet.write(row, col, lineNum)
                        if("P." in t):
                            worksheet.write(row, col + 1, "Patient")
                        else:
                            worksheet.write(row, col + 1, "Doctor") 
                        worksheet.write(row, col + 2, str(firstPart))
                        t1 = (str(firstPartTg)).replace(" ","")
                        t1 = t1.lower()
                        worksheet.write(row, col + 3,  t1)
                        row+=1
                        n+=1
                    if (not ("" == seconedPart)):
                        worksheet.write(row, col, lineNum)
                        if("P." in t):
                            worksheet.write(row, col + 1, "Patient")
                        else:
                            worksheet.write(row, col + 1, "Doctor") 
                        worksheet.write(row, col + 2, str(seconedPart))
                        t1 = (str(seconedPartTg)).replace(" ","")
                        t1 = t1.lower()
                        worksheet.write(row, col + 3,  t1)
                        row+=1
                        n+=1
                    if (not ("" == thirdPart)):
                        worksheet.write(row, col, lineNum)
                        if("P." in t):
                            worksheet.write(row, col + 1, "Patient")
                        else:
                            worksheet.write(row, col + 1, "Doctor") 
                        worksheet.write(row, col + 2, str(thirdPart))
                        t1 = (str(thirdPartTg)).replace(" ","")
                        t1 = t1.lower()
                        worksheet.write(row, col + 3,  t1)
                        row+=1
                        n+=1  
               
                else:
                    worksheet.write(row, col + 2, tempSen)
                    t1 = tg.replace(" ","")
                    t1 = t1.lower()
                    worksheet.write(row, col + 3,  t1)
                    row+=1
                    n+=1
            lau = False
        elif ("Tape" in t): 
            tapeindexes[tapenum]=n
            tapenum+=1
  

#parseWordtoXExcel('C:\hi\DvirCheck.docx')
random.seed(644)

dict = {'personal':1,'laughs':2,'approve':3,'comp':4,'agree':5,'bc':6,'emp':7
        ,'concern':8,'r/o':9,'legit':10,'partner':11,'self-dis':12,'disagree':13,'crit':14,'?reassure':15
        ,'trans':16,'orient':17,'checks':18,'?bid':19,'?understand':20,'?opinion':21,'[?]med':22,'[?]thera':23
        ,'[?]l/s':24,'[?]p/s':25,'[?]other':26,'?med':27,'?thera':28,'?l/s':29,'?p/s':30,'?other':31,'gives-med':32
        ,'gives-thera':33,'gives-l/s':34,'gives-p/s':35,'gives-other':36,'c-med/thera':37,'c-l/s-p/s':38,'?service':39
        ,'unintel':40,'?permission':41}  
hebrewDict = {}

path = "C:\hi\hebrew" +"\\"
path+= str(19) +"_Final.xlsx"
workbook =xlrd.open_workbook(path)
sheet = workbook.sheet_by_name('Sheet1')
hebrewTags= []
englishKeys = []
for key in dict:
    englishKeys.append(key)
    
for rownum in range(sheet.nrows):
        hebrewTags.append(str((sheet.cell_value(rownum, 0))))
i = 0
for item in hebrewTags:
    hebrewDict[item] = englishKeys[i]
    i+=1


train = []
actors = []
tapeindexes[0] = 0
for num in range (14,28):
    path = "C:\hi\hebrew" +"\\"
    path+= str(num) +"_Final.xlsx"
    print(path)
    train_data = []
    train_tags =[]
    actorsData = []
    workbook =xlrd.open_workbook(path)
    if num <=18:
        sheet = workbook.sheet_by_name('Sheet1')
    else:
        sheet = workbook.sheet_by_name('גיליון1')
    print (str(sheet.nrows))
    for rownum in range(sheet.nrows):
        train_data.append(str((sheet.cell_value(rownum, 2))))
    print (train_data)
    for rownum in range(sheet.nrows):
        train_tags.append(str((sheet.cell_value(rownum, 3))))
        
    for rownum in range(sheet.nrows):
        actorsData.append(str((sheet.cell_value(rownum, 1))))
    actorsData=  actorsData[1:]
    train_data = train_data[1:]
    train_tags = train_tags[1:]
    train_tags = list(map(lambda x:hebrewDict[x] ,train_tags))
    ind = 0
    while (ind<len(train_data)):
        train.append((train_data[ind],train_tags[ind]))
        ind+=1
    print (ind)
    if num == 14:
        tapeindexes[1] = ind
    elif num <= 26 and num>=15:
        tapeindexes[num-13]= ind + tapeindexes[num-14]
 
    ind2 = 0
    while (ind2<len(actorsData)):
        actors.append(actorsData[ind2])
        ind2+=1
    #train = train[1:]
actorsData = actors
print(tapeindexes)
#print ("nnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnn ",len(train))
#################################################### parsing ##############################################################################

'''
if isRandom:
    random.shuffle(train)
    train_data = train[:int(0.7*len(train))]
    test_data = train[int(0.7*len(train)):]
    actorsTrain = actorsData[:int(0.7*len(actorsData))]
    actorsTest = actorsData[int(0.7*len(actorsData)):]
else:
    train_data = train[:tapeindexes[4]]
    test_data = train[tapeindexes[4]:]
    actorsTrain = actorsData[:tapeindexes[4]]
    actorsTest = actorsData[tapeindexes[4]:]
'''
    
#################################################### sentences per category #################################################################
    
categoiesSentences = [[] for b in range(41)]    
app = [0 for c in range(41)]
i=0
allCategories = []
for cat in train: 
    t1 = cat[1].replace(" " ,"")
    t1 = t1.lower()    
    allCategories.append(t1)
allCategories = allCategories[1:]    
    
while i<len(train):
    t1 = train[i][1].replace(" " ,"")
    t1 = t1.lower()
    if t1 not in dict:
         print ("sentenceError: " + str(train[i][0]))
         print ("categoryError: " + str(t1))
    else:
        place = int(dict[t1])-1
        app[place]+= 1
        categoiesSentences[place].append(train[i][0])
    i=i+1

row = 1
col = 0

workbook2 =xlsxwriter.Workbook('C:\hi\categoriesSentences.xlsx')
worksheet2 = workbook2.add_worksheet()

for cat in dict:
    worksheet2.write(row, 5, cat)
    row+=1
    for sen in categoiesSentences[dict[cat]-1]:
            worksheet2.write(row, 0, sen)
            row+=1
            
workbook2.close()
#################################################### sentences per category #################################################################


#################################################### into bag of words and features #################################################################
'''
tokenizer1 = keras.preprocessing.text.Tokenizer(filters='!"#$%&*+,.:;<=>@\^_`{|}~\t\n')
actorsTrainVecs = np.vectorize(lambda x: 0 if x == 'רופא:' else 1)(np.array(actorsTrain)).reshape((-1,1))
actorsTestVecs = np.vectorize(lambda x: 0 if x == 'רופא:' else 1)(np.array(actorsTest)).reshape((-1,1))
tokenizer1.fit_on_texts([i[0] for i in train_data])
tokenizer1.fit_on_texts([i[0] for i in test_data])
afterFitText = tokenizer1.texts_to_matrix([i[0] for i in train_data], mode="tfidf")
tokenizer2 = keras.preprocessing.text.Tokenizer(filters='!"#$%&*+,.:;<=>@\^_`{|}~\t\n')
tokenizer2.fit_on_texts([i[1] for i in train_data])
tokenizer2.fit_on_texts([i[1] for i in test_data])
afterFitLabels = tokenizer2.texts_to_matrix([i[1] for i in train_data])
afterFitTextTest = tokenizer1.texts_to_matrix([i[0] for i in test_data], mode="tfidf")
afterFitLabelsTest = tokenizer2.texts_to_matrix([i[1] for i in test_data])

containsQuesionmarkText = np.vectorize(lambda x: 0 if '?' in x else 1)(np.array([i[0] for i in train_data])).reshape((-1,1))
contains3DotsText = np.vectorize(lambda x: 0 if '...' in x else 1)(np.array([i[0] for i in train_data])).reshape((-1,1)) 
contains2DotsText = np.vectorize(lambda x: 0 if '..' in x else 1)(np.array([i[0] for i in train_data])).reshape((-1,1)) 
contains1DotText = np.vectorize(lambda x: 0 if '.' in x else 1)(np.array([i[0] for i in train_data])).reshape((-1,1)) 
containsLinesText = np.vectorize(lambda x: 0 if '--' in x else 1)(np.array([i[0] for i in train_data])).reshape((-1,1)) 
afterFitText = np.hstack((containsQuesionmarkText,afterFitText))
afterFitText = np.hstack((contains3DotsText,afterFitText))
afterFitText = np.hstack((contains2DotsText,afterFitText))
afterFitText = np.hstack((contains1DotText,afterFitText))
afterFitText = np.hstack((containsLinesText,afterFitText))
afterFitText = np.hstack((actorsTrainVecs,afterFitText))

containsQuesionmarkTest = np.vectorize(lambda x: 0 if '?' in x else 1)(np.array([i[0] for i in test_data])).reshape((-1,1))
contains3DotsTest = np.vectorize(lambda x: 0 if '...' in x else 1)(np.array([i[0] for i in test_data])).reshape((-1,1)) 
contains2DotsTest = np.vectorize(lambda x: 0 if '..' in x else 1)(np.array([i[0] for i in test_data])).reshape((-1,1)) 
contains1DotTest = np.vectorize(lambda x: 0 if '.' in x else 1)(np.array([i[0] for i in test_data])).reshape((-1,1)) 
containsLinesTest = np.vectorize(lambda x: 0 if '--' in x else 1)(np.array([i[0] for i in test_data])).reshape((-1,1)) 
afterFitTextTest = np.hstack((containsQuesionmarkTest,afterFitTextTest))
afterFitTextTest = np.hstack((contains3DotsTest,afterFitTextTest))
afterFitTextTest = np.hstack((contains2DotsTest,afterFitTextTest))
afterFitTextTest = np.hstack((contains1DotTest,afterFitTextTest))
afterFitTextTest = np.hstack((containsLinesTest,afterFitTextTest))
afterFitTextTest = np.hstack((actorsTestVecs,afterFitTextTest))
'''
#################################################### into bag of words and features #################################################################



#################################################### nueral network model #################################################################
def nueralModel():
    #afterFitText = np.hstack((actorsTrainVecs,afterFitText))  
    model = Sequential([
        Dense(128, activation='relu', input_shape= afterFitText.shape[1:]),     ##binary classifier 
        Dense(afterFitLabels.shape[1], activation='sigmoid'),
    ])
      
    model.compile(optimizer='rmsprop',
                  loss='categorical_crossentropy',
                  metrics=['accuracy'])
    
    model.fit(afterFitText, afterFitLabels, epochs=20, batch_size=32)
    #curModel = model
    print ("nueral network results: ")
    print(model.evaluate(afterFitTextTest,afterFitLabelsTest))
    return model
#################################################### nueral network model #################################################################




#################################################### nueral network with previous model ###################################################
def nueralWithPreviousModel():
    zerosVec = np.zeros(len(afterFitText[1]))
    newSentencesVec = np.vstack((zerosVec,afterFitText))
    twoSentencesVectorTrain = []
    i = 1
    while i < len(newSentencesVec):
        firstVec = newSentencesVec[i-1]
        seconedVec = newSentencesVec[i]
        tmpVec =  np.append(firstVec,seconedVec)
        twoSentencesVectorTrain.append (tmpVec)
        i+=1
        
    zerosVec = np.zeros(len(afterFitTextTest[1]))
    newSentencesVec = np.vstack((zerosVec,afterFitTextTest))
    twoSentencesVectorTest = []
    i = 1
    while i < len(newSentencesVec):
        firstVec = newSentencesVec[i-1]
        seconedVec = newSentencesVec[i]
        tmpVec =  np.append(firstVec,seconedVec)
        twoSentencesVectorTest.append (tmpVec)
        i+=1
        
    twoSentencesVectorTrain = np.asarray(twoSentencesVectorTrain)
    twoSentencesVectorTest = np.asarray(twoSentencesVectorTest)
    
    model3 = Sequential([
         Dense(256, activation='relu', input_shape= twoSentencesVectorTrain.shape[1:]),    
        Dense(afterFitLabels.shape[1], activation='sigmoid'),  # binary classifier 
    ])
    model3.compile(optimizer='rmsprop',
                  loss='categorical_crossentropy',
                  metrics=['accuracy'])
    
    model3.fit(twoSentencesVectorTrain, afterFitLabels, epochs=20, batch_size=32) 
    print ("nueral network with previous sentence results: ")
    print(model3.evaluate(twoSentencesVectorTest,afterFitLabelsTest))
    return model3
#################################################### nueral network with previous model ###################################################




#################################################### random forest model ##################################################################
def randomForestModel():
    #clf = RandomForestClassifier(n_jobs=10,max_depth=500, random_state=100,n_estimators=200)
    clf = RandomForestClassifier(n_estimators=100, min_samples_split=2, random_state=0,max_features='auto',bootstrap=False,n_jobs=10)
    clf.fit(afterFitText, afterFitLabels)
    print ("random forest results: ")
    print (clf.score(afterFitTextTest,afterFitLabelsTest))
    return clf
#################################################### random forest model ##################################################################




#################################################### nueral network with lstm model ##################################################################

def lstmModel():
    X_train = [i[0] for i in train_data]
    X_test = [i[0] for i in test_data]
    train_word_set = {x.lower() for s in X_train for x in s.split(' ')}
    test_word_set = {x.lower() for s in X_test for x in s.split(' ')}
    word_set = train_word_set.union(test_word_set)
    switch_dict = {word:number+1 for number, word in enumerate(word_set)}
    X_train = list(map(lambda x: [switch_dict[w.lower()] for w in x.split(' ')], X_train))
    X_test = list(map(lambda x: [switch_dict[w.lower()] for w in x.split(' ')], X_test))
    max_length = 20
    X_train = sequence.pad_sequences(X_train, maxlen=max_length)
    X_test = sequence.pad_sequences(X_test, maxlen=max_length)
    top_words = len(X_train)
    max_review_length = 1295
    top_words = len(word_set)+1
    np.random.seed(7)
    max_review_length = 20
    embedding_vecor_length = 32
    
    
    model2 = Sequential()
    model2.add(Embedding(top_words, embedding_vecor_length, input_length=max_review_length))
    model2.add(LSTM(100))
    model2.add(Dense(afterFitLabels.shape[1], activation='softmax'))
    #model2.add(Dense(afterFitLabels.shape[1], activation='sigmoid'))   #binary classifier 
    model2.compile(loss='categorical_crossentropy', optimizer='adam', metrics=['accuracy'])
    model2.fit(X_train, afterFitLabels, epochs=20, batch_size=64)
    
    scores = model2.evaluate(X_test, afterFitLabelsTest, verbose=0)
    print ("lstm results: ")
    print("Accuracy: %.2f%%" % (scores[1]*100))
    return model2

#################################################### nueral network with lstm model ##################################################################

#################################################### k-fold with nueral model ##################################################################
def writeToSheet(workSheet, trasctipt_ID, testSize,trainSize, compName,isComp, speaker, predictiion,real,difference,numCorrect):
    global compositeRow
    workSheet.write(compositeRow, 0, trasctipt_ID)
    workSheet.write(compositeRow, 1, testSize)
    workSheet.write(compositeRow, 2,  trainSize)
    workSheet.write(compositeRow, 3,  compName)
    workSheet.write(compositeRow, 4,  isComp)
    workSheet.write(compositeRow, 5,  speaker)
    workSheet.write(compositeRow, 6,  predictiion)
    workSheet.write(compositeRow, 7,  real)
    workSheet.write(compositeRow, 8,  difference)
    workSheet.write(compositeRow, 9,  numCorrect)
    compositeRow+=1
    
model_predicts  = None
tokenizer_dic = None
tmpTestDataTags = None    
def kFoldNueralNetwork(f,alpha):
    global model_predicts
    global tokenizer_dic
    global tmpTestDataTags
    true_predictions_ratio = 0
    unknown_predictions_ratio= 0 
    #row = 1
    workbook4 =xlsxwriter.Workbook('C:\hi\composites.xlsx')
    worksheet4 = workbook4.add_worksheet()
    worksheet4.write(0, 0, "transcript_ID")
    worksheet4.write(0, 1, "number of utterances in test")
    worksheet4.write(0, 2,  "number of utterances in train")
    worksheet4.write(0, 3,  "composite name/class name")
    worksheet4.write(0, 4,  "is Composite")
    worksheet4.write(0, 5,  "speaker")
    worksheet4.write(0, 6,  "predicted")
    worksheet4.write(0, 7,  "real")
    worksheet4.write(0, 8,  "difference")
    worksheet4.write(0, 9,  "realPredictions")

    

    print (tapeindexes)
    for ind in range(1,15):
        if ind == 1:
            train_data = train[tapeindexes[ind]:]
            test_data = train[:tapeindexes[ind]]
            actorsTrain = actorsData[tapeindexes[ind]:]
            actorsTest = actorsData[:tapeindexes[ind]]
        elif ind == 14:
            train_data = train[:tapeindexes[ind-1]]
            test_data = train[tapeindexes[ind-1]:]
            actorsTrain = actorsData[:tapeindexes[ind-1]]
            actorsTest = actorsData[tapeindexes[ind-1]:]
        else:
            train_data = train[:tapeindexes[ind-1]]
            train_data = train_data + train[tapeindexes[ind]:]
            actorsTrain = actorsData[:tapeindexes[ind-1]]
            actorsTrain = actorsTrain + actorsData[tapeindexes[ind]:]
            test_data = train[tapeindexes[ind-1]:tapeindexes[ind]]
            actorsTest = actorsData[tapeindexes[ind-1]:tapeindexes[ind]]
            
        tmpTestDataTags = [i[1] for i in test_data]
        tokenizer1 = keras.preprocessing.text.Tokenizer(filters='!"#$%&*+,.:;<=>@\^_`{|}~\t\n')
        actorsTrainVecs = np.vectorize(lambda x: 1 if x == 'רופא:' else 0)(np.array(actorsTrain)).reshape((-1,1))
        actorsTestVecs = np.vectorize(lambda x: 1 if x == 'רופא:' else 0)(np.array(actorsTest)).reshape((-1,1))
        tokenizer1.fit_on_texts([i[0] for i in train_data])
        tokenizer1.fit_on_texts([i[0] for i in test_data])
        afterFitText = tokenizer1.texts_to_matrix([i[0] for i in train_data], mode="tfidf")      
        
        #afterFitText = np.hstack((actorsTrainVecs,afterFitText))
        containsQuesionmarkText = np.vectorize(lambda x: 1 if '?' in x else 0)(np.array([i[0] for i in train_data])).reshape((-1,1))
        contains3DotsText = np.vectorize(lambda x: 1 if '...' in x else 0)(np.array([i[0] for i in train_data])).reshape((-1,1)) 
        contains2DotsText = np.vectorize(lambda x: 1 if '..' in x else 0)(np.array([i[0] for i in train_data])).reshape((-1,1)) 
        contains1DotText = np.vectorize(lambda x: 1 if '.' in x else 0)(np.array([i[0] for i in train_data])).reshape((-1,1)) 
        containsLinesText = np.vectorize(lambda x: 1 if '--' in x else 0)(np.array([i[0] for i in train_data])).reshape((-1,1)) 
        afterFitText = np.hstack((containsQuesionmarkText,afterFitText))
        afterFitText = np.hstack((contains3DotsText,afterFitText))
        afterFitText = np.hstack((contains2DotsText,afterFitText))
        afterFitText = np.hstack((contains1DotText,afterFitText))
        afterFitText = np.hstack((containsLinesText,afterFitText))
        #afterFitText = np.hstack((actorsTrainVecs,afterFitText))
        
        
        tokenizer2 = keras.preprocessing.text.Tokenizer(filters='!"#$%&*+,.:;<=>@\^_`{|}~\t\n')
        tokenizer2.fit_on_texts([i[1] for i in train_data])
        tokenizer2.fit_on_texts([i[1] for i in test_data])
        afterFitLabels = tokenizer2.texts_to_matrix([i[1] for i in train_data])
        
        
        ############################################################get Doctor sentences################################
        doctorModels = []
        patientModels = []
        for key in dict:
            model = generateModel(key, afterFitText, train_data,actorsTrain,True)
            doctorModels.append(model)
            model = generateModel(key, afterFitText, train_data,actorsTrain,False)
            patientModels.append(model)
            #K.clear_session()

  
        print ("yeaaaaaaaaaaaaaaaahhhhhhhhhhhhhh") 
        
        afterFitTextTest = tokenizer1.texts_to_matrix([i[0] for i in test_data], mode="tfidf")
        containsQuesionmarkTest = np.vectorize(lambda x: 1 if '?' in x else 0)(np.array([i[0] for i in test_data])).reshape((-1,1))
        contains3DotsTest = np.vectorize(lambda x: 1 if '...' in x else 0)(np.array([i[0] for i in test_data])).reshape((-1,1)) 
        contains2DotsTest = np.vectorize(lambda x: 1 if '..' in x else 0)(np.array([i[0] for i in test_data])).reshape((-1,1)) 
        contains1DotTest = np.vectorize(lambda x: 1 if '.' in x else 0)(np.array([i[0] for i in test_data])).reshape((-1,1)) 
        containsLinesTest = np.vectorize(lambda x: 1 if '--' in x else 0)(np.array([i[0] for i in test_data])).reshape((-1,1)) 
        afterFitTextTest = np.hstack((containsQuesionmarkTest,afterFitTextTest))
        afterFitTextTest = np.hstack((contains3DotsTest,afterFitTextTest))
        afterFitTextTest = np.hstack((contains2DotsTest,afterFitTextTest))
        afterFitTextTest = np.hstack((contains1DotTest,afterFitTextTest))
        afterFitTextTest = np.hstack((containsLinesTest,afterFitTextTest)) 
        #afterFitTextTest = np.hstack((actorsTestVecs,afterFitTextTest))
        
        truePredicts = 0
        unknownPred = 0 
        path = 'C:\hi' +'\\' + 'hebrew_42_Comperation_' + str(ind)+ '.xlsx'
        workbook5 =xlsxwriter.Workbook(path)
        #worksheet2 = workbook2.sheet_by_name('Sheet1')
        worksheet5 = workbook5.add_worksheet()
        worksheet5.write(0, 0, "Sentence")
        worksheet5.write(0, 1, "Real_tag")
        worksheet5.write(0, 2,  "Prediction")
        
        
        for i in range (0,len(afterFitTextTest)):
            if (actorsTestVecs[i] == 1):
                models = doctorModels
            else: 
                models = patientModels
            probs = []
            for model in models:
                probs.append (model.predict(afterFitTextTest[i:i+1]))
            num = np.argmax(probs)
            pred = 'Unknown'
            if (probs[num]>=alpha):
                pred = [key for key, value in dict.items() if value == num+1][0]
            if (pred!='Unknown'):
                if (test_data[i][1] == pred):
                    truePredicts+=1
            else:
                unknownPred+=1
            worksheet5.write(i+1, 0, test_data[i][0])
            worksheet5.write(i+1, 1, test_data[i][1])
            worksheet5.write(i+1, 2, pred)
            
        workbook5.close()
        true_predictions_ratio+= truePredicts/len(afterFitTextTest)
        print ("temp true pretiction ratio:" + str(truePredicts/len(afterFitTextTest)))
        unknown_predictions_ratio+= unknownPred/len(afterFitTextTest)
        print ("temp unknown pretiction ratio:" + str(unknownPred/len(afterFitTextTest)))
        f.flush()
        f.write(str(alpha)+"(tape"+str(ind)+")           "+str(truePredicts/len(afterFitTextTest))+"               "+str(unknownPred/len(afterFitTextTest))+"                 "+str(1-((truePredicts+unknownPred)/len(afterFitTextTest))))
        f.write("\n")
        K.clear_session()

        
    print ("true predictions ration:" + str(true_predictions_ratio/5))
    print ("unknown predictions ration:" + str(unknown_predictions_ratio/5))
    print ("false predictions ration:" + str(1-(true_predictions_ratio/5+unknown_predictions_ratio/5)))
    f.flush()
    f.write(str(alpha)+"           "+str(true_predictions_ratio/5)+"               "+str(unknown_predictions_ratio/5)+"                 "+str(1-(true_predictions_ratio/5+unknown_predictions_ratio/5)))
    f.write("\n")
    f.flush()

    workbook4.close()
    #return predicts, dic
    return None, None
     
     
    #################################################### k-fold with nueral model ##################################################################

    #################################################### comperation between prediction and real category ################################################
def generateModel(key, afterFitText, train_data,actorsTrain,isDoctor):
        modelSentences = np.vectorize(lambda x: 1 if x == key else 0)(np.array([i[1] for i in train_data])).reshape((-1,1))
        if (isDoctor):
            actorsTrainVecs = np.vectorize(lambda x: 1 if x == 'רופא:' else 0)(np.array(actorsTrain)).reshape((-1,1))
        else:
            actorsTrainVecs = np.vectorize(lambda x: 0 if x == 'רופא:' else 1)(np.array(actorsTrain)).reshape((-1,1))
        tmp = np.logical_and(modelSentences , actorsTrainVecs)
        tmp2 = np.vectorize(lambda x: 1 if x == True else 0)(tmp).reshape((-1,1))
        model = Sequential([
            Dense(128, activation='relu', input_shape= afterFitText.shape[1:]),     ##binary classifier 
            Dense(1, activation='sigmoid'),
        ])   
        model.compile(optimizer='rmsprop',
                      loss='binary_crossentropy',
                      metrics=['accuracy'])
        
        model.fit(afterFitText, tmp2, epochs=10, batch_size=32)
        return model

def compareRealAndModel(predicts, afterFitTextTest, tokenizer2, test_data,tapeNum):
    row = 1
    path = 'C:\hi' +'\\' + 'hebrewComperation_' + str(tapeNum)+ '.xlsx'
    workbook2 =xlsxwriter.Workbook(path)
    #worksheet2 = workbook2.sheet_by_name('Sheet1')
    worksheet2 = workbook2.add_worksheet()

    worksheet2.write(0, 0, "Sentence")
    worksheet2.write(0, 1, "Real_tag")
    worksheet2.write(0, 2,  "Prediction")

    
    
    #print (type(model))
    dic = tokenizer2.word_index
    
    tagsNoSpaces =[]
    tagsPredictedNoSpaces =[]
    for i in range(0,len(predicts)):
        p = predicts[i]
        num = np.argmax(p)     
        lKey =[key for key, value in dic.items() if value == num][0]
        worksheet2.write(row, 0, test_data[i][0])
        worksheet2.write(row, 1, test_data[i][1])
        t1 = test_data[i][1].replace(" " ,"")
        t1 = t1.lower()
        tagsNoSpaces.append(t1)
        worksheet2.write(row, 2,  lKey)
        t2 = lKey.replace(" " ,"")
        tagsPredictedNoSpaces.append(t2)
        row+=1

    workbook2.close()
    
def generateConclutions(ind, predicts,tokenizer2,test_data):    
    #compareRealAndModel(model)
    #################################################### comperation between prediction and real category ######################################kf##########
          
    
    
    #################################################### Indices per category ##############################################################################
    #predicts = model.predict(afterFitTextTest)
    dic = tokenizer2.word_index
    row = 1
    
    tagsNoSpaces =[]
    tagsPredictedNoSpaces =[]
    for i in range(0,len(predicts)):
        p = predicts[i]
        num = np.argmax(p)
        lKey =[key for key, value in dic.items() if value == num][0]
        t1 = test_data[row-1][1].replace(" " ,"")
        t1 = t1.lower()
        tagsNoSpaces.append(t1)
        t2 = lKey.replace(" " ,"")
        tagsPredictedNoSpaces.append(t2)
        row+=1
    
    rb = open_workbook('C:\hi\VisualMatrix.xlsx')
    wb = copy(rb) # a writable copy (I can't read values out of this, only write to it)
    w_sheet = wb.get_sheet(6)
    
    tmp = [[0 for c in range(41)] for cc in range(41)] 
    tmp2 = [[0 for b in range(5)] for bb in range(41)] 
    
    for x in range(1, 42):
        for y in range(1, 42):
            w_sheet.write(x, y, int('0'))
    
    index = 0
    
    for item1 in tagsNoSpaces:
        if str(item1) in dict:
            realLoc = int(dict[str(item1)])
        else:
            print ("probbbParsing: " + str(item1))
            continue
                  
        if str(tagsPredictedNoSpaces[index]) in dict:
            preLoc = int(dict[str(tagsPredictedNoSpaces[index])])
        else:
           print ("probbbModel: " +str(tagsPredictedNoSpaces[index]))
           continue
                
        tmp[realLoc-1][preLoc-1]=tmp[realLoc-1][preLoc-1]+1
        if(realLoc == preLoc):
            tmp2[int(realLoc)-1][0]=tmp2[int(realLoc)-1][0]+1 #TP
            tmp2[int(preLoc)-1][4]=tmp2[int(preLoc)-1][4]+1   #TN אמרתי שזה לא זה אבל בפועל זה כן
        else:
            tmp2[int(preLoc)-1][3]=tmp2[int(preLoc)-1][3]+1  #FP
            tmp2[int(realLoc)-1][1]=tmp2[int(realLoc)-1][1]+1 #  #FNאמרתי שכן אבל בפועל זה לא
        tmp2[int(realLoc)-1][2]=tmp2[int(realLoc)-1][3]+tmp2[int(realLoc)-1][0]
        index+=1
    
    for m in range(0, 41):
        for n in range(0, 41):
            w_sheet.write(m+1, n+1, int(str(tmp[m][n])))
            
    for t in range(0, 41):
        for r in range(41, 46):
            w_sheet.write(t+1, r+1, int(str(tmp2[t][r-41])))
            
    for a in range(0, 41):
        if int(str(tmp2[a][0]))+int(str(tmp2[a][3])) ==0:
            w_sheet.write(a+1, 47,"div by 0")
        else:
            w_sheet.write(a+1, 47, float(int(str(tmp2[a][0]))/(int(str(tmp2[a][0]))+int(str(tmp2[a][3])))))
        if int(str(tmp2[a][0]))+int(str(tmp2[a][1])) ==0:
            w_sheet.write(a+1, 48,"div by 0")
        else:
            w_sheet.write(a+1, 48, float(int(str(tmp2[a][0]))/(int(str(tmp2[a][0]))+int(str(tmp2[a][1])))))
        if (int(str(tmp2[a][0]))+int(str(tmp2[a][1])) ==0 or int(str(tmp2[a][0]))+int(str(tmp2[a][3])) ==0
            or int(str(tmp2[a][0]))==0):
            w_sheet.write(a+1, 49,"error")
        else:
            w_sheet.write(a+1, 49,
                          float((2* float(int(str(tmp2[a][0]))/(int(str(tmp2[a][0]))+int(str(tmp2[a][3])))) *float(int(str(tmp2[a][0]))/(int(str(tmp2[a][0]))+int(str(tmp2[a][1])))))/(float(int(str(tmp2[a][0]))/
                                    (int(str(tmp2[a][0]))+int(str(tmp2[a][3])))) +float(int(str(tmp2[a][0]))/(int(str(tmp2[a][0]))+int(str(tmp2[a][1])))))))
    
            
    
    
    wb.save('C:\hi\VisualMatrixNew' +str(ind) + "hebrew" +'.xls')
    #################################################### Indices per category ##############################################################################
    

    #################################################### generating graph ##############################################################################
    
    adjacencyList  = [[] for m in range(41)] 
    tape1= train[:tapeindexes[4]]
    tape1= train
    tape1Tags =  [i[1] for i in test_data]
    index = 0
    while index < len(tape1Tags):
        tag = tape1Tags[index]
        tmpIndex = index+1
        if not adjacencyList[dict[tag]-1]: 
            while tmpIndex< len (tape1Tags):
                adjacencyList[dict[tag]-1].append(dict[tape1Tags[tmpIndex]]-1)
                tmpIndex+=1
        index+=1
    
    i=0
    newAdjacencyList  = [[] for m in range(41)] 
    for lis in adjacencyList:
        sum = 0
        for n in range(41):
            sum+=adjacencyList[i].count(n)
        for n in range(41):
            if (adjacencyList[i].count(n)!=0):
                key = [key for key, value in dict.items() if value == n+1][0]
                newAdjacencyList[i].append((key,adjacencyList[i].count(n)/sum))
        i+=1
    
    
    i=0
    for lis in newAdjacencyList:
        key = [key for key, value in dict.items() if value == i+1][0]
        print(str(key) + ": " + str(lis))
        print()
        i+=1
        
    workbook3 =xlsxwriter.Workbook('C:\hi\graph.xlsx')
    worksheet3 = workbook3.add_worksheet()
    worksheet3.write(0, 0, "From")
    worksheet3.write(0, 1, "To")
    worksheet3.write(0, 2,  "Weight")
    i=0
    row=1
    tmpList = []
    DG=nx.DiGraph()
    for lis in newAdjacencyList:
        key = [key for key, value in dict.items() if value == i+1][0]
        for pair in lis:
                worksheet3.write(row, 0, str(key))
                worksheet3.write(row, 1, str(pair[0]))
                worksheet3.write(row, 2, str(pair[1]))
                tmpList.append((str(key),str(pair[0]),str(pair[1])))
                row+=1
        i+=1
    tmpList.append((str("start"),str(test_data[0][1]),str('0.2')))
    DG.add_weighted_edges_from(tmpList)
    #nx.draw(DG)
    
    elarge = [(u, v) for (u, v, d) in DG.edges(data=True) if float(d['weight']) > 0.2]
    esmall = [(u, v) for (u, v, d) in DG.edges(data=True) if float(d['weight']) <= 0.2]
    pos = nx.spring_layout(DG)  # positions for all nodes
    
# nodes
    nx.draw_networkx_nodes(DG, pos, node_size=600)

# edges
    nx.draw_networkx_edges(DG, pos, edgelist=elarge)
    nx.draw_networkx_edges(DG, pos, edgelist=esmall,
                       alpha=0.5, edge_color='b', style='dashed')

# labels
    nx.draw_networkx_labels(DG, pos, font_size=6, font_family='sans-serif')
    plt.axis('off')
    #plt.show()
    nx.write_pajek(DG, "hebrew_graph" +str(ind) +".net")
    plt.savefig("hebrew_graph"+str(ind)+".png")  
    plt.clf()  
    return newAdjacencyList
    #################################################### generating graph ##############################################################################
        
    #################################################### Composites ##############################################################################
compositeRow = 1
def compositeCorrection (trainTags,predTags):
    sumCorrest = 0
    for f, b in zip(trainTags, predTags):
        if (f == 1):
            if(f == b):
                sumCorrest+=1
    return sumCorrest

def calculateComposites(tagsPredictedNoSpaces,test_data,actorsTestVecs,tapeID,worksheet):
    global compositeRow 
    print ("jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj")
    trainDataWithActors = np.hstack((array(test_data),actorsTestVecs))
    predictionsWithActors = np.hstack((array(tagsPredictedNoSpaces).reshape((-1,1)),actorsTestVecs))
    
    ############################################################# Doctor sentences #####################################################    
    print("Doctor Composites: \n")
    CMEDD = list(map(lambda x: 1 if (str(x[1]) == '[?]med' and int(x[2]) == 0) else 0, trainDataWithActors))
    CMEDD_MODEL = list(map(lambda x: 1 if (str(x[0]) == '[?]med' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(CMEDD,CMEDD_MODEL)
    CMEDD = sum(CMEDD)
    CMEDD_MODEL = sum (CMEDD_MODEL)
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"[?]med","No","Doctor",CMEDD_MODEL,CMEDD, abs(CMEDD_MODEL-CMEDD), correct)
    CTHERD = list(map(lambda x: 1 if (str(x[1]) == '[?]thera' and int(x[2]) == 0) else 0, trainDataWithActors))
    CTHERD_MODEL = list(map(lambda x: 1 if (str(x[0]) == '[?]thera' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(CTHERD,CTHERD_MODEL)
    CTHERD = sum(CTHERD)
    CTHERD_MODEL = sum (CTHERD_MODEL)
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"[?]thera","No","Doctor",CTHERD_MODEL,CTHERD, abs(CTHERD_MODEL-CTHERD),correct)
    COTHD =  list(map(lambda x: 1 if (str(x[1]) == '[?]other' and int(x[2]) == 0) else 0, trainDataWithActors))
    COTHD_MODEL = list(map(lambda x: 1 if (str(x[0]) == '[?]other' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(COTHD,COTHD_MODEL)
    COTHD = sum(COTHD)
    COTHD_MODEL = sum (COTHD_MODEL)
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"[?]other","No","Doctor",COTHD_MODEL,COTHD, abs(COTHD_MODEL-COTHD),correct)
    OMEDD = list(map(lambda x: 1 if (str(x[1]) == '?med' and int(x[2]) == 0) else 0, trainDataWithActors))
    OMEDD_MODEL = list(map(lambda x: 1 if (str(x[0]) == '?med' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(OMEDD,OMEDD_MODEL)
    OMEDD = sum(OMEDD)
    OMEDD_MODEL = sum (OMEDD_MODEL)
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?med","No","Doctor",OMEDD_MODEL,OMEDD, abs(OMEDD_MODEL-OMEDD),correct)
    OTHERD = list(map(lambda x: 1 if (str(x[1]) == '?thera' and int(x[2]) == 0) else 0, trainDataWithActors))
    OTHERD_MODEL = list(map(lambda x: 1 if (str(x[0]) == '?thera' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(OTHERD,OTHERD_MODEL)
    OTHERD = sum(OTHERD)
    OTHERD_MODEL = sum (OTHERD_MODEL)
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?thera","No","Doctor",OTHERD_MODEL,OTHERD, abs(OTHERD_MODEL-OTHERD),correct)
    OOTHD = list(map(lambda x: 1 if (str(x[1]) == '?other' and int(x[2]) == 0) else 0, trainDataWithActors))
    OOTHD_MODEL = list(map(lambda x: 1 if (str(x[0]) == '?other' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(OOTHD,OOTHD_MODEL)
    OOTHD = sum(OOTHD)
    OOTHD_MODEL = sum (OOTHD_MODEL)
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?other","No","Doctor",OOTHD_MODEL,OOTHD, abs(OOTHD_MODEL-OOTHD),correct)
    BIDD = list(map(lambda x: 1 if (str(x[1]) == '?bid' and int(x[2]) == 0) else 0, trainDataWithActors))
    BIDD_MODEL = list(map(lambda x: 1 if (str(x[0]) == '?bid' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(BIDD,BIDD_MODEL)
    BIDD = sum(BIDD)
    BIDD_MODEL = sum (BIDD_MODEL)
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?bid","No","Doctor",BIDD_MODEL,BIDD, abs(BIDD_MODEL-BIDD),correct)

    MEDQUED = CMEDD+CTHERD+COTHD+OMEDD+OTHERD+OOTHD+BIDD
    MEDQUED_MODEL = CMEDD_MODEL+CTHERD_MODEL+COTHD_MODEL+OMEDD_MODEL+OTHERD_MODEL+OOTHD_MODEL+BIDD_MODEL 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"MEDQUED","Yes","Doctor",MEDQUED_MODEL,MEDQUED, abs(MEDQUED_MODEL-MEDQUED),"none")
    print ("MEDQUED: " , abs(MEDQUED-MEDQUED_MODEL), MEDQUED, MEDQUED_MODEL)

    CLSD = list(map(lambda x: 1 if (str(x[1]) == '[?]l/s' and int(x[2]) == 0) else 0, trainDataWithActors))
    CLSD_MODEL = list(map(lambda x: 1 if (str(x[0]) == '[?]l/s' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(CLSD,CLSD_MODEL)
    CLSD = sum(CLSD)
    CLSD_MODEL = sum (CLSD_MODEL)
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"[?]l/s","No","Doctor",CLSD_MODEL,CLSD, abs(CLSD_MODEL-CLSD),correct)
    CPSD = list(map(lambda x: 1 if (str(x[1]) == '[?]p/s' and int(x[2]) == 0) else 0, trainDataWithActors))
    CPSD_MODEL = list(map(lambda x: 1 if (str(x[0]) == '[?]p/s' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(CPSD,CPSD_MODEL)
    CPSD = sum(CPSD)
    CPSD_MODEL = sum (CPSD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"[?]p/s","No","Doctor",CPSD_MODEL,CPSD, abs(CPSD_MODEL-CPSD),correct)
    OLSD = list(map(lambda x: 1 if (str(x[1]) == '?l/s' and int(x[2]) == 0) else 0, trainDataWithActors))
    OLSD_MODEL = list(map(lambda x: 1 if (str(x[0]) == '?l/s' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(OLSD,OLSD_MODEL)
    OLSD = sum(OLSD)
    OLSD_MODEL = sum (OLSD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?l/s","No","Doctor",OLSD_MODEL,OLSD, abs(OLSD_MODEL-OLSD),correct)
    OPSD = list(map(lambda x: 1 if (str(x[1]) == '?p/s' and int(x[2]) == 0) else 0, trainDataWithActors))
    OPSD_MODEL = list(map(lambda x: 1 if (str(x[0]) == '?p/s' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(OPSD,OPSD_MODEL)
    OPSD = sum(OPSD)
    OPSD_MODEL = sum (OPSD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?p/s","No","Doctor",OPSD_MODEL,OPSD, abs(OPSD_MODEL-OPSD),correct)

    PSYQUED=CLSD+CPSD+OLSD+OPSD
    PSYQUED_MODEL = CLSD_MODEL+CPSD_MODEL+OLSD_MODEL+OPSD_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"PSYQUED","Yes","Doctor",PSYQUED_MODEL,PSYQUED, abs(PSYQUED_MODEL-PSYQUED),"none")
    print ("PSYQUED :" , abs (PSYQUED-PSYQUED_MODEL), PSYQUED, PSYQUED_MODEL)
    
    IMEDD = list(map(lambda x: 1 if (str(x[1]) == 'gives-med' and int(x[2]) == 0) else 0, trainDataWithActors))
    IMEDD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'gives-med' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(IMEDD,IMEDD_MODEL)
    IMEDD = sum(IMEDD)
    IMEDD_MODEL = sum (IMEDD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"gives-med","No","Doctor",IMEDD_MODEL,IMEDD, abs(IMEDD_MODEL-IMEDD),correct)
    ITHERD = list(map(lambda x: 1 if (str(x[1]) == 'gives-thera' and int(x[2]) == 0) else 0, trainDataWithActors))
    ITHERD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'gives-thera' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(ITHERD,ITHERD_MODEL)
    ITHERD = sum(ITHERD)
    ITHERD_MODEL = sum (ITHERD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"gives-thera","No","Doctor",ITHERD_MODEL,ITHERD, abs(ITHERD_MODEL-ITHERD),correct)
    IOTHD = list(map(lambda x: 1 if (str(x[1]) == 'gives-other' and int(x[2]) == 0) else 0, trainDataWithActors))
    IOTHD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'gives-other' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(IOTHD,IOTHD_MODEL)
    IOTHD = sum(IOTHD)
    IOTHD_MODEL = sum (IOTHD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"gives-other","No","Doctor",IOTHD_MODEL,IOTHD, abs(IOTHD_MODEL-IOTHD),correct)
    CNLMDD = list(map(lambda x: 1 if (str(x[1]) == 'c-med/thera' and int(x[2]) == 0) else 0, trainDataWithActors))
    CNLMDD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'c-med/thera' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(CNLMDD,CNLMDD_MODEL)
    CNLMDD = sum(CNLMDD)
    CNLMDD_MODEL = sum (CNLMDD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"c/med-thera","No","Doctor",CNLMDD_MODEL,CNLMDD, abs(CNLMDD_MODEL-CNLMDD),correct)

    INFOMEDD=IMEDD+ITHERD+IOTHD+CNLMDD
    INFOMEDD_MODEL=IMEDD_MODEL+ITHERD_MODEL+IOTHD_MODEL+CNLMDD_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"INFOMEDD","Yes","Doctor",INFOMEDD_MODEL,INFOMEDD, abs(INFOMEDD_MODEL-INFOMEDD),"none")
    print ("INFOMEDD :" , abs (INFOMEDD-INFOMEDD_MODEL), INFOMEDD,INFOMEDD_MODEL)
    
    
    ILSD = list(map(lambda x: 1 if (str(x[1]) == 'gives-l/s' and int(x[2]) == 0) else 0, trainDataWithActors))
    ILSD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'gives-l/s' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(ILSD,ILSD_MODEL)
    ILSD = sum(ILSD)
    ILSD_MODEL = sum (ILSD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"gives-l/s","No","Doctor",ILSD_MODEL,ILSD, abs(ILSD_MODEL-ILSD),correct)
    IPSD = list(map(lambda x: 1 if (str(x[1]) == 'gives-p/s' and int(x[2]) == 0) else 0, trainDataWithActors))
    IPSD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'gives-p/s' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(IPSD,IPSD_MODEL)
    IPSD = sum(IPSD)
    IPSD_MODEL = sum (IPSD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"gives-p/s","No","Doctor",IPSD_MODEL,IPSD, abs(IPSD_MODEL-IPSD),correct)
    CNLLSD = list(map(lambda x: 1 if (str(x[1]) == 'c-l/s-p/s' and int(x[2]) == 0) else 0, trainDataWithActors))
    CNLLSD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'c-l/s-p/s' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(CNLLSD,CNLLSD_MODEL)
    CNLLSD = sum(CNLLSD)
    CNLLSD_MODEL = sum (CNLLSD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"c-l/s-p/s","No","Doctor",CNLLSD_MODEL,CNLLSD, abs(CNLLSD_MODEL-CNLLSD),correct)

    INFOPSYD=ILSD+IPSD+CNLLSD
    INFOPSYD_MODEL=ILSD_MODEL+IPSD_MODEL+CNLLSD_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"INFOPSYD","Yes","Doctor",INFOPSYD_MODEL,INFOPSYD, abs(INFOPSYD_MODEL-INFOPSYD),"none")
    print ("INFOPSYD :" , abs (INFOPSYD-INFOPSYD_MODEL), INFOPSYD-INFOPSYD_MODEL)


    ASKOD = list(map(lambda x: 1 if (str(x[1]) == '?opinion' and int(x[2]) == 0) else 0, trainDataWithActors))
    ASKOD_MODEL = list(map(lambda x: 1 if (str(x[0]) == '?opinion' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(ASKOD,ASKOD_MODEL)
    ASKOD = sum(ASKOD)
    ASKOD_MODEL = sum (ASKOD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?opinion","No","Doctor",ASKOD_MODEL,ASKOD, abs(ASKOD_MODEL-ASKOD),correct)
    ASKPD =  list(map(lambda x: 1 if (str(x[1]) == '?permission' and int(x[2]) == 0) else 0, trainDataWithActors))
    ASKPD_MODEL = list(map(lambda x: 1 if (str(x[0]) == '?permission' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(ASKPD,ASKPD_MODEL)
    ASKPD = sum(ASKPD)
    ASKPD_MODEL = sum (ASKPD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?permission","No","Doctor",ASKPD_MODEL,ASKPD, abs(ASKPD_MODEL-ASKPD),correct)
    ASKRD = list(map(lambda x: 1 if (str(x[1]) == '?reassure' and int(x[2]) == 0) else 0, trainDataWithActors))
    ASKRD_MODEL = list(map(lambda x: 1 if (str(x[0]) == '?reassure' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(ASKRD,ASKRD_MODEL)
    ASKRD = sum(ASKRD)
    ASKRD_MODEL = sum (ASKRD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?reassure","No","Doctor",ASKRD_MODEL,ASKRD, abs(ASKRD_MODEL-ASKRD),correct)
    ASKUD = list(map(lambda x: 1 if (str(x[1]) == '?understand' and int(x[2]) == 0) else 0, trainDataWithActors))
    ASKUD_MODEL = list(map(lambda x: 1 if (str(x[0]) == '?understand' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(ASKUD,ASKUD_MODEL)
    ASKUD = sum(ASKUD)
    ASKUD_MODEL = sum (ASKUD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?understand","No","Doctor",ASKUD_MODEL,ASKUD, abs(ASKUD_MODEL-ASKUD),correct)
    BCD = list(map(lambda x: 1 if (str(x[1]) == 'bc' and int(x[2]) == 0) else 0, trainDataWithActors))
    BCD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'bc' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(BCD,BCD_MODEL)
    BCD = sum(BCD)
    BCD_MODEL = sum (BCD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"bc","No","Doctor",BCD_MODEL,BCD, abs(BCD_MODEL-BCD),correct)
    CHECD = list(map(lambda x: 1 if (str(x[1]) == 'checks' and int(x[2]) == 0) else 0, trainDataWithActors))
    CHECD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'checks' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(CHECD,CHECD_MODEL)
    CHECD = sum(CHECD)
    CHECD_MODEL = sum (CHECD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"checks","No","Doctor",CHECD_MODEL,CHECD, abs(CHECD_MODEL-CHECD),correct)

    PARTNERD=ASKOD+ASKPD+ASKRD+ASKUD+BCD+CHECD
    PARTNERD_MODEL=ASKOD_MODEL+ASKPD_MODEL+ASKRD_MODEL+ASKUD_MODEL+BCD_MODEL+CHECD_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"PARTNERD","Yes","Doctor",PARTNERD_MODEL,PARTNERD, abs(PARTNERD_MODEL-PARTNERD),"none")
    print ("PARTNERD :" , abs (PARTNERD-PARTNERD_MODEL), PARTNERD,PARTNERD_MODEL)


    LAUGD = list(map(lambda x: 1 if (str(x[1]) == 'laughs' and int(x[2]) == 0) else 0, trainDataWithActors))
    LAUGD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'laughs' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(LAUGD,LAUGD_MODEL)
    LAUGD = sum(LAUGD)
    LAUGD_MODEL = sum (LAUGD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"laughs","No","Doctor",LAUGD_MODEL,LAUGD, abs(LAUGD_MODEL-LAUGD),correct)
    APPD = list(map(lambda x: 1 if (str(x[1]) == 'approve' and int(x[2]) == 0) else 0, trainDataWithActors))
    APPD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'approve' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(APPD,APPD_MODEL)
    APPD = sum(APPD)
    APPD_MODEL = sum (APPD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"approve","No","Doctor",APPD_MODEL,APPD, abs(APPD_MODEL-APPD),correct)
    COMPD = list(map(lambda x: 1 if (str(x[0]) == 'comp' and int(x[1]) == 0) else 0, predictionsWithActors))
    COMPD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'comp' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(COMPD,COMPD_MODEL)
    COMPD = sum(COMPD)
    COMPD_MODEL = sum (COMPD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"comp","No","Doctor",COMPD_MODEL,COMPD, abs(COMPD_MODEL-COMPD),correct)
    AGRED = list(map(lambda x: 1 if (str(x[1]) == 'agree' and int(x[2]) == 0) else 0, trainDataWithActors))
    AGRED_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'agree' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(AGRED,AGRED_MODEL)
    AGRED = sum(AGRED)
    AGRED_MODEL = sum (AGRED_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"agree","No","Doctor",AGRED_MODEL,AGRED, abs(AGRED_MODEL-AGRED),correct)

    POSD=LAUGD+APPD+COMPD+AGRED
    POSD_MODEL=LAUGD_MODEL+APPD_MODEL+COMPD_MODEL+AGRED_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"POSD","Yes","Doctor",POSD_MODEL,POSD, abs(POSD_MODEL-POSD),"none")
    print ("POSD :" , abs (POSD-POSD_MODEL), POSD,POSD_MODEL)


    EMPD =  list(map(lambda x: 1 if (str(x[1]) == 'emp' and int(x[2]) == 0) else 0, trainDataWithActors))
    EMPD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'emp' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(EMPD,EMPD_MODEL)
    EMPD = sum(EMPD)
    EMPD_MODEL = sum (EMPD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"emp","No","Doctor",EMPD_MODEL,EMPD, abs(EMPD_MODEL-EMPD),correct)
    LEGITD = list(map(lambda x: 1 if (str(x[1]) == 'legit' and int(x[2]) == 0) else 0, trainDataWithActors))
    LEGITD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'legit' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(LEGITD,LEGITD_MODEL)
    LEGITD = sum(LEGITD)
    LEGITD_MODEL = sum (LEGITD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"legit","No","Doctor",LEGITD_MODEL,LEGITD, abs(LEGITD_MODEL-LEGITD),correct)
    COND = list(map(lambda x: 1 if (str(x[1]) == 'concern' and int(x[2]) == 0) else 0, trainDataWithActors))
    COND_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'concern' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(COND,COND_MODEL)
    COND = sum(COND)
    COND_MODEL = sum (COND_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"concern","No","Doctor",COND_MODEL,COND, abs(COND_MODEL-COND),correct)
    ROD = list(map(lambda x: 1 if (str(x[1]) == 'r/o' and int(x[2]) == 0) else 0, trainDataWithActors))
    ROD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'r/o' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(ROD,ROD_MODEL)
    ROD = sum(ROD)
    ROD_MODEL = sum (ROD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"r/o","No","Doctor",ROD_MODEL,ROD, abs(ROD_MODEL-ROD),correct)
    PARTD = list(map(lambda x: 1 if (str(x[1]) == 'partner' and int(x[2]) == 0) else 0, trainDataWithActors))
    PARTD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'partner' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(PARTD,PARTD_MODEL)
    PARTD = sum(PARTD)
    PARTD_MODEL = sum (PARTD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"partner","No","Doctor",PARTD_MODEL,PARTD, abs(PARTD_MODEL-PARTD),correct)
    SDISD = list(map(lambda x: 1 if (str(x[1]) == 'self-dis' and int(x[2]) == 0) else 0, trainDataWithActors))
    SDISD_MODEL =list(map(lambda x: 1 if (str(x[0]) == 'self-dis' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(SDISD,SDISD_MODEL)
    SDISD = sum(SDISD)
    SDISD_MODEL = sum (SDISD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"self-dis","No","Doctor",SDISD_MODEL,SDISD, abs(SDISD_MODEL-SDISD),correct)

    EMOD=EMPD+LEGITD+COND+ROD+PARTD+SDISD
    EMOD_MODEL=EMPD_MODEL+LEGITD_MODEL+COND_MODEL+ROD_MODEL+PARTD_MODEL+SDISD_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"EMOD","Yes","Doctor",EMOD_MODEL,EMOD, abs(EMOD_MODEL-EMOD),"none")
    print ("EMOD :" , abs (EMOD-EMOD_MODEL), EMOD,EMOD_MODEL)


    DISD = list(map(lambda x: 1 if (str(x[1]) == 'disagree' and int(x[2]) == 0) else 0, trainDataWithActors))
    DISD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'disagree' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(DISD,DISD_MODEL)
    DISD = sum(DISD)
    DISD_MODEL = sum (DISD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"disagree","No","Doctor",DISD_MODEL,DISD, abs(DISD_MODEL-DISD),correct)
    CRITD = list(map(lambda x: 1 if (str(x[1]) == 'crit' and int(x[2]) == 0) else 0, trainDataWithActors))
    CRITD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'crit' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(CRITD,CRITD_MODEL)
    CRITD = sum(CRITD)
    CRITD_MODEL = sum (CRITD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"crit","No","Doctor",CRITD_MODEL,CRITD, abs(CRITD_MODEL-CRITD),correct)

    NEGD=DISD+CRITD
    NEGD_MODEL=DISD_MODEL+CRITD_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"NEGD","Yes","Doctor",NEGD_MODEL,NEGD, abs(NEGD_MODEL-NEGD),correct)
    print ("NEGD :" , abs (NEGD-NEGD_MODEL), NEGD,NEGD_MODEL)


    PERSD = list(map(lambda x: 1 if (str(x[1]) == 'personal' and int(x[2]) == 0) else 0, trainDataWithActors))
    PERSD_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'personal' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(PERSD,PERSD_MODEL)
    PERSD = sum(PERSD)
    PERSD_MODEL = sum (PERSD_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"personal","No","Doctor",PERSD_MODEL,PERSD, abs(PERSD_MODEL-PERSD),correct)

    CHITD=PERSD
    CHITD_MODEL = PERSD_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"CHITD","Yes","Doctor",CHITD_MODEL,CHITD, abs(CHITD_MODEL-CHITD),"none")
    print ("CHITD :" , abs (CHITD-CHITD_MODEL), CHITD,CHITD_MODEL)


    TRAND = list(map(lambda x: 1 if (str(x[1]) == 'trans' and int(x[2]) == 0) else 0, trainDataWithActors))
    TRAND_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'trans' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(TRAND,TRAND_MODEL)
    TRAND = sum(TRAND)
    TRAND_MODEL = sum (TRAND_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"trans","No","Doctor",TRAND_MODEL,TRAND, abs(TRAND_MODEL-TRAND),correct)
    ORID = list(map(lambda x: 1 if (str(x[1]) == 'orient' and int(x[2]) == 0) else 0, trainDataWithActors))
    ORID_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'orient' and int(x[1]) == 0) else 0, predictionsWithActors))
    correct = compositeCorrection(ORID,ORID_MODEL)
    ORID = sum(ORID)
    ORID_MODEL = sum (ORID_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"orient","No","Doctor",ORID_MODEL,ORID, abs(ORID_MODEL-ORID),correct)

    PROCD=TRAND+ORID
    PROCD_MODEL=TRAND_MODEL+ORID_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"PROCD","Yes","Doctor",PROCD_MODEL,PROCD, abs(PROCD_MODEL-PROCD),"none")
    print ("PROCD :" , abs (PROCD-PROCD_MODEL), PROCD,PROCD_MODEL)

    ############################################################# Doctor sentences #####################################################
    
    ############################################################# Patient sentences #####################################################
    print("\nPatient Composites: \n")
    QMEDP = list(map(lambda x: 1 if ((str(x[1]) == '?med' or str(x[1]) == '[?]med') and int(x[2]) == 1) else 0, trainDataWithActors))
    QMEDP_MODEL = list(map(lambda x: 1 if ((str(x[0]) == '?med' or str(x[0]) == '[?]med') and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(QMEDP,QMEDP_MODEL)
    QMEDP = sum(QMEDP)
    QMEDP_MODEL = sum (QMEDP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?med and [?]med","No","Patient",QMEDP_MODEL,QMEDP, abs(QMEDP_MODEL-QMEDP),correct)
    QTHERP = list(map(lambda x: 1 if ((str(x[1]) == '?thera' or str(x[1]) == '[?]thera') and int(x[2]) == 1) else 0, trainDataWithActors))
    QTHERP_MODEL = list(map(lambda x: 1 if ((str(x[0]) == '?thera' or str(x[0]) == '[?]thera') and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(QTHERP,QTHERP_MODEL)
    QTHERP = sum(QTHERP)
    QTHERP_MODEL = sum (QTHERP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?thera and [?]thera","No","Patient",QTHERP_MODEL,QTHERP, abs(QTHERP_MODEL-QTHERP),correct)
    QOTHP = list(map(lambda x: 1 if ((str(x[1]) == '?other' or str(x[1]) == '[?]other') and int(x[2]) == 1) else 0, trainDataWithActors))
    QOTHP_MODEL = list(map(lambda x: 1 if ((str(x[0]) == '?other' or str(x[0]) == '[?]other') and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(QOTHP,QOTHP_MODEL)
    QOTHP = sum(QOTHP)
    QOTHP_MODEL = sum (QOTHP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?other and [?]other","No","Patient",QOTHP_MODEL,QOTHP, abs(QOTHP_MODEL-QOTHP),correct)
    BIDP = list(map(lambda x: 1 if (str(x[1]) == '?bid' and int(x[2]) == 1) else 0, trainDataWithActors))
    BIDP_MODEL = list(map(lambda x: 1 if (str(x[0]) == '?bid' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(BIDP,BIDP_MODEL)
    BIDP = sum(BIDP)
    BIDP_MODEL = sum (BIDP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?bid","No","Patient",BIDP_MODEL,BIDP, abs(BIDP_MODEL-BIDP),correct)

    MEDQUEP=QMEDP+QTHERP+QOTHP+BIDP
    MEDQUEP_MODEL=QMEDP_MODEL+QTHERP_MODEL+QOTHP_MODEL+BIDP_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"MEDQUEP","Yes","Patient",MEDQUEP_MODEL,MEDQUEP, abs(MEDQUEP_MODEL-MEDQUEP),"none")
    print ("MEDQUEP :" , abs (MEDQUEP-MEDQUEP_MODEL), MEDQUEP,MEDQUEP_MODEL)

    QLSP = list(map(lambda x: 1 if ((str(x[1]) == '?l/s' or str(x[1]) == '[?]l/s') and int(x[2]) == 1) else 0, trainDataWithActors))
    QLSP_MODEL = list(map(lambda x: 1 if ((str(x[0]) == '?l/s' or str(x[0]) == '[?]l/s') and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(QLSP,QLSP_MODEL)
    QLSP = sum(QLSP)
    QLSP_MODEL = sum (QLSP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?l/s","No","Patient",QLSP_MODEL,QLSP, abs(QLSP_MODEL-QLSP),correct)
    QPSP = list(map(lambda x: 1 if ((str(x[1]) == '?p/s' or str(x[1]) == '[?]p/s') and int(x[2]) == 1) else 0, trainDataWithActors))
    QPSP_MODEL = list(map(lambda x: 1 if ((str(x[0]) == '?p/s' or str(x[0]) == '[?]p/s') and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(QPSP,QPSP_MODEL)
    QPSP = sum(QPSP)
    QPSP_MODEL = sum (QPSP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?p/s","No","Patient",QPSP_MODEL,QPSP, abs(QPSP_MODEL-QPSP),correct)

    PSYQUEP=QLSP+QPSP
    PSYQUEP_MODEL=QLSP_MODEL+QPSP_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"PSYQUEP","Yes","Patient",PSYQUEP_MODEL,PSYQUEP, abs(PSYQUEP_MODEL-PSYQUEP),"none")
    print ("PSYQUEP :" , abs (PSYQUEP-PSYQUEP_MODEL), PSYQUEP,PSYQUEP_MODEL)


    IMEDP = list(map(lambda x: 1 if (str(x[1]) == 'gives-med' and int(x[2]) == 1) else 0, trainDataWithActors))
    IMEDP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'gives-med' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(IMEDP,IMEDP_MODEL)
    IMEDP = sum(IMEDP)
    IMEDP_MODEL = sum (IMEDP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"gives-med","No","Patient",IMEDP_MODEL,IMEDP, abs(IMEDP_MODEL-IMEDP),correct)
    ITHERP = list(map(lambda x: 1 if (str(x[1]) == 'gives-thera' and int(x[2]) == 1) else 0, trainDataWithActors))
    ITHERP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'gives-thera' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(ITHERP,ITHERP_MODEL)
    ITHERP = sum(ITHERP)
    ITHERP_MODEL = sum (ITHERP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"gives-thera","No","Patient",ITHERP_MODEL,ITHERP, abs(ITHERP_MODEL-ITHERP),correct)
    IOTHP = list(map(lambda x: 1 if (str(x[1]) == 'gives-other' and int(x[2]) == 1) else 0, trainDataWithActors))
    IOTHP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'gives-other' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(IOTHP,IOTHP_MODEL)
    IOTHP = sum(IOTHP)
    IOTHP_MODEL = sum (IOTHP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"gives-other","No","Patient",IOTHP_MODEL,IOTHP, abs(IOTHP_MODEL-IOTHP),correct)

    INFOMEDP=IMEDP+ITHERP+IOTHP
    INFOMEDP_MODEL=IMEDP_MODEL+ITHERP_MODEL+IOTHP_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"INFOMEDP","Yes","Patient",INFOMEDP_MODEL,INFOMEDP, abs(INFOMEDP_MODEL-INFOMEDP),"none")
    print ("INFOMEDP :" , abs (INFOMEDP-INFOMEDP_MODEL), INFOMEDP,INFOMEDP_MODEL)

    ILSP = list(map(lambda x: 1 if (str(x[1]) == 'gives-l/s' and int(x[2]) == 1) else 0, trainDataWithActors))
    ILSP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'gives-l/s' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(ILSP,ILSP_MODEL)
    ILSP = sum(ILSP)
    ILSP_MODEL = sum (ILSP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"gives-l/s","No","Patient",ILSP_MODEL,ILSP, abs(ILSP_MODEL-ILSP),correct)
    IPSP = list(map(lambda x: 1 if (str(x[1]) == 'gives-p/s' and int(x[2]) == 1) else 0, trainDataWithActors))
    IPSP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'gives-p/s' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(IPSP,IPSP_MODEL)
    IPSP = sum(IPSP)
    IPSP_MODEL = sum (IPSP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"gives-p/s","No","Patient",IPSP_MODEL,IPSP, abs(IPSP_MODEL-IPSP),correct)

    INFOPSYP=ILSP+IPSP
    INFOPSYP_MODEL=ILSP_MODEL+IPSP_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"INFOPSYP","Yes","Patient",INFOPSYP_MODEL,INFOPSYP, abs(INFOPSYP_MODEL-INFOPSYP),"none")
    print ("INFOPSYP :" , abs (INFOPSYP-INFOPSYP_MODEL),INFOPSYP,INFOPSYP_MODEL)


    ASKSP = list(map(lambda x: 1 if (str(x[1]) == '?service' and int(x[2]) == 1) else 0, trainDataWithActors))
    ASKSP_MODEL = list(map(lambda x: 1 if (str(x[0]) == '?service' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(ASKSP,ASKSP_MODEL)
    ASKSP = sum(ASKSP)
    ASKSP_MODEL = sum (ASKSP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?service","No","Patient",ASKSP_MODEL,ASKSP, abs(ASKSP_MODEL-ASKSP),correct)
    ASKRP = list(map(lambda x: 1 if (str(x[1]) == '?reassure' and int(x[2]) == 1) else 0, trainDataWithActors))
    ASKRP_MODEL = list(map(lambda x: 1 if (str(x[0]) == '?reassure' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(ASKRP,ASKRP_MODEL)
    ASKRP = sum(ASKRP)
    ASKRP_MODEL = sum (ASKRP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?reassure","No","Patient",ASKRP_MODEL,ASKRP, abs(ASKRP_MODEL-ASKRP),correct)
    ASKUP = list(map(lambda x: 1 if (str(x[1]) == '?understand' and int(x[2]) == 1) else 0, trainDataWithActors))
    ASKUP_MODEL = list(map(lambda x: 1 if (str(x[0]) == '?understand' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(ASKUP,ASKUP_MODEL)
    ASKUP = sum(ASKUP)
    ASKUP_MODEL = sum (ASKUP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"?understand","No","Patient",ASKUP_MODEL,ASKUP, abs(ASKUP_MODEL-ASKUP),correct)
    CHECP = list(map(lambda x: 1 if (str(x[1]) == 'checks' and int(x[2]) == 1) else 0, trainDataWithActors))
    CHECP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'checks' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(CHECP,CHECP_MODEL)
    CHECP = sum(CHECP)
    CHECP_MODEL = sum (CHECP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"checks","No","Patient",CHECP_MODEL,CHECP, abs(CHECP_MODEL-CHECP),correct)

    PARTNERP=ASKSP+ASKRP+ASKUP+CHECP
    PARTNERP_MODEL=ASKSP_MODEL+ASKRP_MODEL+ASKUP_MODEL+CHECP_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"PARTNERP","Yes","Patient",PARTNERP_MODEL,PARTNERP, abs(PARTNERP_MODEL-PARTNERP),"none")
    print ("PARTNERP :" , abs (PARTNERP-PARTNERP_MODEL), PARTNERP,PARTNERP_MODEL)

    LAUGP = list(map(lambda x: 1 if (str(x[1]) == 'laughs' and int(x[2]) == 1) else 0, trainDataWithActors))
    LAUGP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'laughs' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(LAUGP,LAUGP_MODEL)
    LAUGP = sum(LAUGP)
    LAUGP_MODEL = sum (LAUGP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"laughs","No","Patient",LAUGP_MODEL,LAUGP, abs(LAUGP_MODEL-LAUGP),correct)
    APPP = list(map(lambda x: 1 if (str(x[1]) == 'approve' and int(x[2]) == 1) else 0, trainDataWithActors))
    APPP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'approve' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(APPP,APPP_MODEL)
    APPP = sum(APPP)
    APPP_MODEL = sum (APPP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"approve","No","Patient",APPP_MODEL,APPP, abs(APPP_MODEL-APPP),correct)
    COMPP = list(map(lambda x: 1 if (str(x[1]) == 'comp' and int(x[2]) == 1) else 0, trainDataWithActors))
    COMPP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'comp' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(COMPP,COMPP_MODEL)
    COMPP = sum(COMPP)
    COMPP_MODEL = sum (COMPP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"comp","No","Patient",COMPP_MODEL,COMPP, abs(COMPP_MODEL-COMPP),correct)
    AGREP = list(map(lambda x: 1 if (str(x[1]) == 'agree' and int(x[2]) == 1) else 0, trainDataWithActors))
    AGREP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'agree' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(AGREP,AGREP_MODEL)
    AGREP = sum(AGREP)
    AGREP_MODEL = sum (AGREP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"agree","No","Patient",AGREP_MODEL,AGREP, abs(AGREP_MODEL-AGREP),correct)

    POSP=LAUGP+APPP+COMPP+AGREP 
    POSP_MODEL=LAUGP_MODEL+APPP_MODEL+COMPP_MODEL+AGREP_MODEL 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"POSP","Yes","Patient",POSP_MODEL,POSP, abs(POSP_MODEL-POSP),"none")
    print ("POSP :" , abs (POSP-POSP_MODEL), POSP,POSP_MODEL)

    EMPP = list(map(lambda x: 1 if (str(x[1]) == 'emp' and int(x[2]) == 1) else 0, trainDataWithActors))
    EMPP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'emp' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(EMPP,EMPP_MODEL)
    EMPP = sum(EMPP)
    EMPP_MODEL = sum (EMPP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"emp","No","Patient",EMPP_MODEL,EMPP, abs(EMPP_MODEL-EMPP),correct)
    LEGITP = list(map(lambda x: 1 if (str(x[1]) == 'legit' and int(x[2]) == 1) else 0, trainDataWithActors))
    LEGITP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'legit' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(LEGITP,LEGITP_MODEL)
    LEGITP = sum(LEGITP)
    LEGITP_MODEL = sum (LEGITP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"legit","No","Patient",LEGITP_MODEL,LEGITP, abs(LEGITP_MODEL-LEGITP),correct)
    CONP = list(map(lambda x: 1 if (str(x[1]) == 'concern' and int(x[2]) == 1) else 0, trainDataWithActors))
    CONP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'concern' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(CONP,CONP_MODEL)
    CONP = sum(CONP)
    CONP_MODEL = sum (CONP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"concern","No","Patient",CONP_MODEL,CONP, abs(CONP_MODEL-CONP),correct)
    ROP = list(map(lambda x: 1 if (str(x[1]) == 'r/o' and int(x[2]) == 1) else 0, trainDataWithActors))
    ROP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'r/o' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(ROP,ROP_MODEL)
    ROP = sum(ROP)
    ROP_MODEL = sum (ROP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"r/o","No","Patient",ROP_MODEL,ROP, abs(ROP_MODEL-ROP),correct)

    EMOP=EMPP+LEGITP+CONP+ROP
    EMOP_MODEL=EMPP_MODEL+LEGITP_MODEL+CONP_MODEL+ROP_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"EMOP","Yes","Patient",EMOP_MODEL,EMOP, abs(EMOP_MODEL-EMOP),"none")
    print ("EMOP :" , abs (EMOP-EMOP_MODEL), EMOP,EMOP_MODEL)
    
    DISP = list(map(lambda x: 1 if (str(x[1]) == 'disagree' and int(x[2]) == 1) else 0, trainDataWithActors))
    DISP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'disagree' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(DISP,DISP_MODEL)
    DISP = sum(DISP)
    DISP_MODEL = sum (DISP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"disagree","No","Patient",DISP_MODEL,DISP, abs(DISP_MODEL-DISP),correct)
    CRITP = list(map(lambda x: 1 if (str(x[1]) == 'crit' and int(x[2]) == 1) else 0, trainDataWithActors))
    CRITP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'crit' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(CRITP,CRITP_MODEL)
    CRITP = sum(CRITP)
    CRITP_MODEL = sum (CRITP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"crit","No","Patient",CRITP_MODEL,CRITP, abs(CRITP_MODEL-CRITP),correct)

    NEGP=DISP+CRITP
    NEGP_MODEL=DISP_MODEL+CRITP_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"NEGP","Yes","Patient",NEGP_MODEL,NEGP, abs(NEGP_MODEL-NEGP),"none")
    print ("NEGP :" , abs (NEGP-NEGP_MODEL), NEGP, NEGP_MODEL)

    PERSP = list(map(lambda x: 1 if (str(x[1]) == 'personal' and int(x[2]) == 1) else 0, trainDataWithActors))
    PERSP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'personal' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(PERSP,PERSP_MODEL)
    PERSP = sum(PERSP)
    PERSP_MODEL = sum (PERSP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"personal","No","Patient",PERSP_MODEL,PERSP, abs(PERSP_MODEL-PERSP),correct)
   
    CHITP=PERSP
    CHITP_MODEL=PERSP_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"CHITP","Yes","Patient",CHITP_MODEL,CHITP, abs(CHITP_MODEL-CHITP),"none")
    print ("CHITP :" , abs (CHITP-CHITP_MODEL), CHITP,CHITP_MODEL)


    TRANP = list(map(lambda x: 1 if (str(x[1]) == 'trans' and int(x[2]) == 1) else 0, trainDataWithActors))
    TRANP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'trans' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(TRANP,TRANP_MODEL)
    TRANP = sum(TRANP)
    TRANP_MODEL = sum (TRANP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"trans","No","Patient",TRANP_MODEL,TRANP, abs(TRANP_MODEL-TRANP),correct)
    ORIP = list(map(lambda x: 1 if (str(x[1]) == 'orient' and int(x[2]) == 1) else 0, trainDataWithActors))
    ORIP_MODEL = list(map(lambda x: 1 if (str(x[0]) == 'orient' and int(x[1]) == 1) else 0, predictionsWithActors))
    correct = compositeCorrection(ORIP,ORIP_MODEL)
    ORIP = sum(ORIP)
    ORIP_MODEL = sum (ORIP_MODEL) 
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"orient","No","Patient",ORIP_MODEL,ORIP, abs(ORIP_MODEL-ORIP),correct)

    PROCP=TRANP+ORIP
    PROCP_MODEL=TRANP_MODEL+ORIP_MODEL
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"PROCP","Yes","Patient",PROCP_MODEL,PROCP, abs(PROCP_MODEL-PROCP),"none")
    print ("PROCP :" , abs (PROCP-PROCP_MODEL), PROCP,PROCP_MODEL)
    
    ############################################################# Patient sentences ####################################################################
    print("\nMixed Composites: \n")

    
    compute_ptcent1 = (PSYQUED + INFOPSYD + EMOD + PSYQUEP + PARTNERD + INFOPSYP + EMOP +  MEDQUEP  )/ ( MEDQUED + PROCD + INFOMEDP + INFOMEDD)
    compute_ptcent1_MODEL = (PSYQUED_MODEL + INFOPSYD_MODEL + EMOD_MODEL + PSYQUEP_MODEL + PARTNERD_MODEL + INFOPSYP_MODEL + EMOP_MODEL +  MEDQUEP_MODEL )/ ( MEDQUED_MODEL + PROCD_MODEL + INFOMEDP_MODEL + INFOMEDD_MODEL)
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"compute_ptcent1","Yes","Both",compute_ptcent1_MODEL,compute_ptcent1, abs(compute_ptcent1_MODEL-compute_ptcent1),"none")
    print ("compute_ptcent1 :" , abs (compute_ptcent1-compute_ptcent1_MODEL), compute_ptcent1,compute_ptcent1_MODEL)

    compute_ptcent2 = (PSYQUED + INFOPSYD + EMOD + PSYQUEP + PARTNERD + INFOPSYP + EMOP + MEDQUEP + INFOMEDD )/ ( MEDQUED + PROCD + INFOMEDP )
    compute_ptcent2_MODEL = (PSYQUED_MODEL + INFOPSYD_MODEL + EMOD_MODEL + PSYQUEP_MODEL + PARTNERD_MODEL + INFOPSYP_MODEL + EMOP_MODEL + MEDQUEP_MODEL + INFOMEDD_MODEL )/ ( MEDQUED_MODEL + PROCD_MODEL + INFOMEDP_MODEL )
    writeToSheet(worksheet,tapeID,len(test_data),1316-len(test_data),"compute_ptcent2","Yes","Both",compute_ptcent2_MODEL,compute_ptcent2, abs(compute_ptcent2_MODEL-compute_ptcent2),"none")
    print ("compute_ptcent2 :" , abs (compute_ptcent2-compute_ptcent2_MODEL),compute_ptcent2,compute_ptcent2_MODEL)
    #################################################### Composites ##############################################################################
  
    #################################################### Composites ##############################################################################
def predictionsUsingGraph(graph):
    b = [[0 for c in range(len(tokenizer_dic))] for cc in range(len(model_predicts))]
    for m in range (0,len(tokenizer_dic)):
        b[0][m] = model_predicts[0][m]
    for i in range (1,len(model_predicts)):
        for j in range (0,len(tokenizer_dic)):
                    p = model_predicts[i-1]
                    num = np.argmax(p)
                    nowCat = [key for key, value in tokenizer_dic.items() if value == j+1][0]
                    prevCat = [key for key, value in tokenizer_dic.items() if value == num+1][0]
                    place = dict[prevCat]
                    #tmplis = graph[place]
                    lis = graph[place]
                    #lKey =[key for key, value in tokenizer_dic.items() if value == num][0]
                    prevProb = 1
                    for pair in lis:
                        if  str(pair[0]) == str(nowCat):
                            prevProb = pair[1]     
                    b[i][j] = model_predicts[i][j]*prevProb
    #print (b)
    i=0
    newPredictions = []
    for i in range (0,len(model_predicts)):
        p = b[i]
        num = np.argmax(p)
        nowCat = [key for key, value in tokenizer_dic.items() if value == num+1][0]
        newPredictions.append(nowCat)
        #i+=1
    print (newPredictions)
    i = 0
    score = 0
    for pred in newPredictions:
        if (str(pred) == str(tmpTestDataTags[i])):
            score+=1
        i+=1
    print (score/len(newPredictions))
    return b
        
    

def main():
    #afterFitTextTest,tokenizer2 = kFoldNueralNetwork()
   # generateConclutions(afterFitTextTest,test_data,tokenizer2,tmpModel)
    f = open('file.txt','w')
    f.write("alpha           true_predictions_ratio         unknown_predictions_ratio                 false_predictions_ratio")
    f.write("\n")
    f.flush()
    print ("jjjjjjjjjjjjjjjjj")
    for alpha in np.arange(0,0.4,0.05):    
        model_predicts, tokenizer_dic = kFoldNueralNetwork(f,alpha)
    
    
if __name__ == "__main__":
    main()
