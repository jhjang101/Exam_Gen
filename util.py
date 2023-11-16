# import
import os
from zipfile import ZipFile
from bs4 import BeautifulSoup
import re
import pandas as pd
import random
import numpy as np


# filename and list of forms
def file_in():
    """
    returns original file name and the list of the form.
    Ask a original file name and number of forms to generate.
    Append original and save file name to a list.
    """
    # ask original file name
    print("type filename without file extension. (Ex. test)")
    file = input()

    # ask munber of form to generate
    print("How many forms? (min=2, max=8)")
    how_many = int(input())

    # make list of form
    forms = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
    forms = forms[:how_many]

    return(file, forms)


# read docx and parse xml
def read_docx(file):
    """
    Makes soup from the original docx file using ZipFile and BeautifulSoup.
    :param file: file name of the original docx without extension. 
    """
    # open docx with zipfile
    with ZipFile(file + ".docx", 'r') as zip:
        doc = zip.read('word/document.xml')

    # parse xml 
    soup = BeautifulSoup(doc, 'xml')

    # # print lines
    # paragraphs = soup.find_all('w:p')  # retrieves all paragraphs from the soup
    # for i, paragraph in enumerate(paragraphs):
    #     line = paragraph.get_text()
    #     print(i, line)

    # # print soup
    # print(soup.prettify())

    return(soup)


# read questions and keys 
def read_questions(soup):
    """
    Generate dataframe containing questions, answers, and keys from a BeautifulSoup object, soup.
    :param soup: a BeautifulSoup object contains xml of the original docx. 
    """
    
    paragraphs = soup.find_all('w:p')

    df = pd.DataFrame(columns=['Q','Q_text',
                               'A','A_text',
                               'B','B_text',
                               'C','C_text',
                               'D','D_text',
                               'K','K_text'])
    # count the number of questions, answers, and keys for error
    No_of_Questions = 0
    No_of_A = 0
    No_of_B = 0
    No_of_C = 0
    No_of_D = 0
    No_of_Keys = 0
    key_start = 0
    
    for i, paragraph in enumerate(paragraphs):
        text = paragraph.get_text()
        if bool(re.match(r'\d+\. ', text)):    # extract questions
            df.loc[len(df), ['Q','Q_text']] = [i, text]
            
            No_of_Questions += 1    # count the number of questions
            
        elif bool(re.match(r'A\.', text)):    # extract answers A
            df.loc[len(df)-1, ['A','A_text']] = [i, text]

            No_of_A += 1    # count number of answers A
            
        elif bool(re.match(r'B\.', text)):    # extract answers B
            df.loc[len(df)-1, ['B','B_text']] = [i, text]

            No_of_B += 1    # count number of answers B
            
        elif bool(re.match(r'C\.', text)):    # extract answers C
            df.loc[len(df)-1, ['C','C_text']] = [i, text]

            No_of_C += 1    # count number of answers C
            
        elif bool(re.match(r'D\.', text)):    # extract answers D
            df.loc[len(df)-1, ['D','D_text']] = [i, text]
            
            No_of_D += 1    # count number of answers D
            
        elif bool(re.match('FORM X KEY', text)):
            key_start = i+1    # find the position that starts the keys
            
        else:
            pass
    
    # show message if key is not detected
    if key_start==0:
        print('FORM X KEY is not detected')

    else:
        pass


    # extract keys
    for i in range(No_of_Questions):
        K = key_start+i    # the position of the key
        K_text = paragraphs[K].get_text()    # the key (ex. A B C D)
        if bool(re.match('A|B|C|D', K_text)):
            df.loc[i, ['K','K_text']] = [K, K_text]
            
            No_of_Keys += 1    # count number of keys
        
        else:
            pass
        
    print('Number of questions: ', No_of_Questions)
    print('Number of A: ', No_of_A)
    print('Number of B: ', No_of_B)
    print('Number of C: ', No_of_C)
    print('Number of D: ', No_of_D)
    print('Number of keys: ', No_of_Keys)
    print('')

    # show message if number of question and number of key is not same
    if No_of_Questions == No_of_Keys:
        pass
    else:
        print('The number of answers is not the same as the number of questions.')
        print(df)
        
    # create df_key with question number, original key
    df_key = pd.DataFrame()
    df_key['Q'] = list(range(1,len(df)+1))
    df_key['Original'] = df['K_text']
    
    return(df, df_key)


# shuffle the answers
def shuffle(df):
    """
    Generate a shuffled dataframe containing questions, shuffled answers, and new keys from df.
    :param df: a BeautifulSoup object contains xml of the original docx. 
    """
    # generate df_to_shuffle
    df_to_shuffle= df[['Q','A','B','C','D','K','K_text']].dropna()

    # copy df_to_shuffle to df_shuffled
    df_shuffled = df_to_shuffle.copy()
    
    # list of the key (ex. [A, B, C, D])
    K_texts = df_to_shuffle['K_text'].tolist()

    # list of position of the correct answers
    K_indexes = []
    
    # lookup the position of the correct answers and add to the K_index column
    for i, K_text in enumerate(K_texts):
        K_index = df_to_shuffle[K_text].tolist()[i]
        K_indexes.append(K_index)
    df_shuffled['K_index']=K_indexes
    
    # set the K_index column type as integer
    df_shuffled['K_index']=df_shuffled['K_index'].astype('Int64')
    
    # shuffle the answers and assign new key for each questions
    new_K_texts = []
    for i, K_index in enumerate(K_indexes):
        answer = df_shuffled.iloc[i,1:5].tolist()    # list of answer position
        random.shuffle(answer)    # shuffle
        df_shuffled.iloc[i,1:5]=answer    # update shuffled answer to the df_shuffled
        new_K_text = df_shuffled.columns[df_shuffled.iloc[i].isin([K_index])][0]    # lookup the new key (ABCD) after shuffle
        new_K_texts.append(new_K_text)
    df_shuffled['K_text'] = new_K_texts    # update new key (ABCD)

    print(f'{len(df_shuffled)} out of {len(df)} questions are shuffled.')
    
    return(df_to_shuffle,df_shuffled)


# update list of paragraphs
def to_paragraphs(soup, df_to_shuffle, df_shuffled, letter ='A'):
    """
    1. Reorder paragraphs position according to the shuffled dataframe.
    2. Replace the FORM X to A, B, C, D, ...
    :param soup: a BeautifulSoup object contains xml of the original docx. 
    :param df_to_shuffle: a dataframe before shuffle
    :param df_shuffled: a dataframe after shuffle
    :param letter: a string of alphabet letter for form number
    """
    # get paragraphs from dummy and replace
    paragraphs = soup.find_all('w:p')
    paragraphs_dummy = paragraphs.copy()
 
    ###### for A
    new_As = df_shuffled['A'].tolist()    # list of new A position
    As = df_to_shuffle['A'].tolist()    # list of original A position

    for i, n in enumerate(new_As):
        paragraphs[As[i]] = paragraphs_dummy[n]    # replace A paragraph

    # change ABCD to A
    for n in As:
        A_ttag = paragraphs[n].find('w:t')
        A_text = A_ttag.get_text()  
        new_A_text = 'A.' + A_text[2:]   
        A_ttag.string = new_A_text

    ###### for B
    new_Bs = df_shuffled['B'].tolist()    # list of new B position
    Bs = df_to_shuffle['B'].tolist()    # list of original B position

    for i, n in enumerate(new_Bs):
        paragraphs[Bs[i]] = paragraphs_dummy[n]    # replace B paragraph

    # change ABCD to B
    for n in Bs:
        B_ttag = paragraphs[n].find('w:t')
        B_text = B_ttag.get_text()  
        new_B_text = 'B.' + B_text[2:]   
        B_ttag.string = new_B_text
    
    ###### for C
    new_Cs = df_shuffled['C'].tolist()    # list of new C position
    Cs = df_to_shuffle['C'].tolist()    # list of original C position

    for i, n in enumerate(new_Cs):
        paragraphs[Cs[i]] = paragraphs_dummy[n]    # replace C paragraph

    # change ABCD to B
    for n in Cs:
        C_ttag = paragraphs[n].find('w:t')
        C_text = C_ttag.get_text()  
        new_C_text = 'C.' + C_text[2:]   
        C_ttag.string = new_C_text
    
    ###### for D    
    new_Ds = df_shuffled['D'].tolist()    # list of new D position
    Ds = df_to_shuffle['D'].tolist()    # list of original D position
    
    for i, n in enumerate(new_Ds):
        paragraphs[Ds[i]] = paragraphs_dummy[n]    # replace D paragraph
        
    # change ABCD to D
    for n in Ds:
        D_ttag = paragraphs[n].find('w:t')
        D_text = D_ttag.get_text()  
        new_D_text = 'D.' + D_text[2:]   
        D_ttag.string = new_D_text
    
    # replace form X
    for i, paragraph in enumerate(paragraphs):
        line = paragraph.get_text()
        if bool(re.search('FORM:', line)):    # find paragraph that has FORM: X
            texts = paragraph.find_all('w:t')
            texts[-1].string = letter

    return(paragraphs)


# build xml
def to_xml(soup, paragraphs):
    """
    re-build xml. Get head (before body) and tail (after body including margin) from the original docx.
    remove key section
    :param soup: a BeautifulSoup object contains xml of the original docx. 
    :param paragraphs: reordered paragraphs after shuffle.
    """
    
    # get string of xml head
    head_pos = str(soup).find('<w:body>')+8    # find a position of head from soup
    head = str(soup)[:head_pos]
    
    # find paragraph position right before keys section
    for i, p in enumerate(paragraphs):    
        string = p.get_text()
        if bool(re.search('FORM X KEY', string[:10])):
            last = i-1
    
    # remove the key section
    questions_section__pos_list = paragraphs[:last]

    # join the paragraph and convert to string
    questions_section = ''.join(str(l) for l in questions_section__pos_list)

    # find a position of tail from soup
    tail_pos = str(soup).rfind('<w:sectPr')
    tail = str(soup)[tail_pos:]

    # create soup_new by joining head, questions_section, and tail
    xml =  head + questions_section + tail
    
    return(xml)


# update key table
def update_keytable(df_key,df_shuffled,letter):
    """
    Add a new column containing new key to the df_key
    :df_key: dataframe containing question number and original key
    :df_shuffled: dataframe containing new keys colume
    :param letter: a string of alphabet letter for form number
    """    
    # add new colume on df_key
    df_key[f'Form {letter}'] = df_key['Original']
    
    # update the new colume with shuffled key
    df_newkey = df_shuffled[['K_text']]
    df_newkey.columns=[f'Form {letter}']
    df_key.update(df_newkey)

    return(df_key)


# save docx and csv
def to_file(xml,df_key,file,letter):
    """
    1. creates a copy of the docx file without 'word/document.xml' in a new folder.
    2. add document.xml from xml 
    :param xml: docx xml string rebuilt from to_xml.
    :param file: file name of the original docx without extension.
    :param letter: a string of alphabet letter for form number
    """
    # save file name
    fileout = file+'_'+letter+'.docx'
    filepath = os.path.join(f'{file}/',fileout)    

    # make a folder
    os.makedirs(file,exist_ok=True)
    
    # creates a copy of the docx file without 'word/document.xml'
    with ZipFile(file + '.docx', 'r') as zipin:
        with ZipFile(filepath, 'w') as zipout:
            zipout.comment = zipin.comment # preserve the comment
            for item in zipin.infolist():
                if item.filename != 'word/document.xml':
                    zipout.writestr(item, zipin.read(item.filename))
    
    
    # add "word/document.xml' containing xml
    with ZipFile(filepath, 'a') as newzip:
        newzip.writestr('word/document.xml', xml)
        print(fileout, 'is added to', file, 'folder.\n')
        
    # save the key file name
    keyout = file+'_key.csv'
    keypath = os.path.join(f'{file}/',keyout)  
    
    # make a folder
    os.makedirs(file,exist_ok=True)
    
    # save
    df_key.to_csv(keypath, index=False)