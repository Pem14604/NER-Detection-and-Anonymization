import os
import pandas as pd
import nltk
from nltk.tag import StanfordNERTagger
from nltk.tokenize import word_tokenize
nltk.download('stopwords')
from nltk.corpus import stopwords
from faker import Faker
fake = Faker()
import win32api
import win32com.client
import spacy
import re
nlp = spacy.load("en_core_web_sm")

java_path = r'C:\Program Files\Java\jdk-12.0.1\bin\java.exe'
os.environ['JAVAHOME'] = java_path
os.environ['CLASSPATH'] = r'C:\Users\vsolanki\AppData\Local\Programs\Python\Python37\Lib\site-packages\stanford-ner-2015-04-20\stanford-ner.jar'
os.environ['STANFORD_MODELS']=r'C:\Users\vsolanki\AppData\Local\Programs\Python\Python37\Lib\site-packages\stanford-ner-2015-04-20\classifiers'
stanford_classifier=r'C:\Users\vsolanki\AppData\Local\Programs\Python\Python37\Lib\site-packages\stanford-ner-2015-04-20\classifiers\english.all.3class.distsim.crf.ser.gz'
st = StanfordNERTagger(stanford_classifier)


document_path_folder = 'C:\\Users\\vsolanki\\Documents\\'
document_file_name = 'rahul_bio.docx'
document_path = document_path_folder + document_file_name
f=open('Indian_name.TXT','r')
Indian_name = f.read()


def doc_to_text(document_path):
    import win32com.client
    doc = win32com.client.GetObject(document_path)
    text = doc.Range().Text
    text = text.replace("\t", " ")
    text = text.replace("\r", " ")
    text = text.replace("!@#$%^&*()[]{};:,./<>?\|`~-=_+", " ")
    text = text.replace("\x07", " ")
    text = text.replace("\x0b", " ")
    text = text.replace("\xa0", " ")
    text = text.replace("\x0c", " ")
    text = text.replace("\x01", " ")
    #text = text.lower()
    #print (text)
    return text


def NER_Extration(text):
    import spacy
    import re
    email = re.findall(r'[\w\.-]+@[\w\.-]+', text)
    date = re.findall(r'\d{4}-\d{2}-\d{2}', text)
    date1= re.findall(r'\d{2}-\d{2}-\d{4}', text)
    year= re.findall(r'\d{4}', text)
    phone_number = re.findall(r'\d{10}', text)
    #phone_number = re.findall(r'[\+\(]?[1-9][0-9.\-\(\)]{8,}[0-9]', text)
    # examples can extract numbers in the formats ['+60 (0)3 2723 7900', '+60 (0)3 2723 7900', '60 (0)4 255 9000', '+6 (03) 8924 8686', '+6 (03) 8924 8000', '60 (7) 268-6200', '+60 (7) 228-6202', '+601-4228-8055']
    nlp = spacy.load("en_core_web_sm")
    doc = nlp(text)
    Named_entityORG=[]
    Named_entityGPE=[]
    Named_entityPERSON=[]
    for ent in doc.ents:
        if (ent.label_ =="ORG"):
              Named_entityORG.append(ent.text)
        if (ent.label_ =="GPE"):
              Named_entityGPE.append(ent.text)
        if (ent.label_ =="PERSON"):
              Named_entityPERSON.append(ent.text)
    tokenized_text = word_tokenize(text)
    classified_text = st.tag(tokenized_text)
    American_Names=[]
    for name,tag in classified_text :
        if tag == 'PERSON':
            American_Names.append(name)
    Inidan_Names=[]
    for word in text.split(): 
        if word in Indian_name.split():
            Inidan_Names.append(word)
    Named_Entity =Named_entityPERSON + American_Names + Inidan_Names
    return email,date1, phone_number,Named_entityORG,Named_entityGPE,Named_Entity,year


def search_replace_all(word_file, find_str, replace_str):
    import win32com.client
    from shutil import copyfile
    ''' replace all occurrences of `find_str` w/ `replace_str` in `word_file` '''
    wdFindContinue = 1
    wdReplaceAll = 2
    # Dispatch() attempts to do a GetObject() before creating a new one.
    # DispatchEx() just creates a new one. 
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = 0
    app.DisplayAlerts = 0
    app.Documents.Open(word_file)
    app.Selection.Find.Execute(find_str, False, False, False, False, False, \
        True, wdFindContinue, False, replace_str, wdReplaceAll)
    app.ActiveDocument.Close(SaveChanges=True)
    app.Quit()


def creating_encrition_doc(document_path_folder, document_file_name):
    from shutil import copyfile
    word_file = copyfile(document_path_folder + document_file_name , document_path_folder + 'encripted_' + document_file_name)
    return word_file


def Anonymization_Doc_File_Data(document_path_folder, document_file_name,document_path):
    text = doc_to_text('C:\\Users\\vsolanki\\Documents\\rahul_bio.docx')
    text = str(text)
    email,date1, phone_number,Named_entityORG,Named_entityGPE,Named_Entity,year= NER_Extration(text)
   
    word_file = creating_encrition_doc(document_path_folder, document_file_name)
    for i in email:
        i = str(i)
        search_replace_all(word_file, i , fake.email())
    for i in date1:
        i = str(i)
        search_replace_all(word_file, i , fake.date()) 
    for i in year:
        i = str(i)
        search_replace_all(word_file, i , fake.year())     
    for i in phone_number:
        i = str(i)
        search_replace_all(word_file, i , fake.phone_number())
    for i in Named_entityORG:
        search_replace_all(word_file, i , fake.city())
        i = str(i)
    for i in Named_entityGPE:
        search_replace_all(word_file, i , fake.city())
    for i in Named_Entity:
        i = str(i)
        search_replace_all(word_file, i , fake.name())

Anonymization_Doc_File_Data(document_path_folder, document_file_name,document_path)



