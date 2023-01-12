from zipfile import ZipFile
from bs4 import BeautifulSoup
import unicodedata
import re
import pandas as pd
import random


###################################
# Function to change xml unicode text to string
###################################
def UnicodeToString(paragraph):
    """
    UnicodeToString nomarize unicode text to string
    :paragraph: element of paragraphs
    """
    line = paragraph.get_text()
    text = unicodedata.normalize('NFKD', line)
    return(text)


###################################
# Create Class Paragraph
###################################
class Paragraph:
    
    # Attribute constructor
    def __init__(self, body):
        self.list = body.find_all('w:p')
    
    
    # Method
    def last(self):
        """
        last method returns the last paragraph index.
        If there is answer key, it retures a paragraph index right before the answer key.
        """
        last = len(self.list)-1    # last paragraph index
        
        # if first 10 charactors in the paragraph has 'FORM A KEY' save the line number-1 to last_line_nbr.
        for i, p in enumerate(self.list):    
            string = UnicodeToString(p)
            if bool(re.search('FORM A KEY', string[:10])):
                last = i-1
        
        return(last)
    
    
    # Method
    def get_questions(self, search='\d+\. ', size=4):
        """
        get_questions method returns a Dictionary contains 'question_index' and 'question'.
        """ 
        # construct empty dictionary 'questions'
        questions = {'Q_index':[],'Q':[]}  
        
        for i, p in enumerate(self.list):
            string = UnicodeToString(p)

            # if first 4 charactors in the paragraph has 'number dot space' save it to questions dict.
            if bool(re.search(search, string[:size])):
                # print(i,string)
                questions['Q_index'].append(i)
                questions['Q'].append(string)
        
        return(pd.DataFrame(questions))
    
    
    # Method
    def get_answers(self, search='A\. ', size=4):
        """
        get_answers method returns a Dictionary contains 'answer_index' and 'answer'.
        """ 
        # construct empty dictionary 'questions'
        answers = {'answer_index':[],'answer':[]}  
        
        for i, p in enumerate(self.list):
            string = UnicodeToString(p)

            # if first 4 charactors in the paragraph has 'number dot space' save it to questions dict.
            if bool(re.search(search, string[:size])):
                # print(i,string)
                answers['answer_index'].append(i)
                answers['answer'].append(string)
        
        return(pd.DataFrame(answers))
    
    
    # Method
    def get_keys(self, search='FORM A KEY', size=10):
        """
        get_keys method returns a Dictionary contains 'keys'.
        """ 
        # construct empty dictionary 'keys'
        keys = {'key_index':[],'key':[]}
        
        # if first 10 charactors in the paragraph has 'FORM A KEY' save the index to key_ln.
        for i, p in enumerate(self.list):
            string = UnicodeToString(p)
            if bool(re.search(search, string[:size])):
                # print(i,string)
                key_ln = i+1
        
        # Save keys to keys dict.
        last = len(self.list)-1    # last paragraph index
        for i in range(key_ln,last):
            p = self.list[i]
            string = UnicodeToString(p)
            keys['key_index'].append(i)
            keys['key'].append(string)
        
        return(pd.DataFrame(keys))


###################################
# Function to append a index to dataframe
###################################
def AppendIndex(questions, index):
    """
    AppendIndex append a index to questions dataframe.
    :questions: dataframe that contains column 'question_index'
    :index: a number. usally paragraph index of EndOfQuestions
    Appending EndOfQuestions index is important for combining questions and answers.
    """
    df_index = pd.DataFrame({'Q_index':index}, index=[0])
    df_append = pd.concat([questions,df_index],ignore_index=True)
    return(df_append)


###################################
# Function to combine Questions, Answers, and Keys
###################################
def JoinQandA(df_q, df_a, df_b, df_c, df_d, df_keys):
    """
    JoinQandA find answers and pair with respected question
    :df: dataframe containing columns 'question_index' and 'question'
    :df_a: dataframe containing columns 'answer_index' and 'answer'
    """

    # construct empty dictionary 'questions'
    df = pd.DataFrame(dict(Q_index=pd.Series([], dtype='Int64'), 
                         Q=pd.Series([], dtype='object'), 
                         A_index=pd.Series([], dtype='Int64'),
                         A=pd.Series([], dtype='object'),
                         B_index=pd.Series([], dtype='Int64'),
                         B=pd.Series([], dtype='object'),
                         C_index=pd.Series([], dtype='Int64'),
                         C=pd.Series([], dtype='object'),
                         D_index=pd.Series([], dtype='Int64'),
                         D=pd.Series([], dtype='object'),
                         key_index=pd.Series([], dtype='Int64'),
                         keys=pd.Series([], dtype='object')
                        ))
    
    ##########    for df_q    ##############
    df['Q_index']=df_q['Q_index']
    df['Q']=df_q['Q']
    
    # find answer index between two question index
    for i in range(len(df_q)-1):
        start=df['Q_index'][i] # search from question index
        end=df['Q_index'][i+1] # search until the next question index
             
        ##########    for df_a    ##############
        # return pd series containing A, and A_index for question 
        bool = df_a['answer_index'].between(start,end)
        A_index = df_a[bool]['answer_index']
        A = df_a[bool]['answer']
        
        # if there is A_index for the question append A_index to df
        if A_index.empty==False:
            df.loc[i,'A_index']=A_index.tolist()[0]
        
        # if there is A for the question append to A to df       
        if A.empty==False:
            df.loc[i,'A']=A.tolist()[0]
    
        ##########    for df_b    ##############        
        # return pd series containing B, and B_index for question 
        bool = df_b['answer_index'].between(start,end)
        B_index = df_b[bool]['answer_index']
        B = df_b[bool]['answer']
        
        # if there is B_index for the question append B_index to df
        if B_index.empty==False:
            df.loc[i,'B_index']=B_index.tolist()[0]
        
        # if there is B for the question append to B to df       
        if B.empty==False:
            df.loc[i,'B']=B.tolist()[0]
            
        ##########    for df_c    ##############        
        # return pd series containing C, and C_index for question 
        bool = df_c['answer_index'].between(start,end)
        C_index = df_c[bool]['answer_index']
        C = df_c[bool]['answer']
        
        # if there is C_index for the question append C_index to df
        if C_index.empty==False:
            df.loc[i,'C_index']=C_index.tolist()[0]
        
        # if there is C for the question append to C to df       
        if C.empty==False:
            df.loc[i,'C']=C.tolist()[0]
            
        ##########    for df_d    ##############        
        # return pd series containing D, and D_index for question 
        bool = df_d['answer_index'].between(start,end)
        D_index = df_d[bool]['answer_index']
        D = df_d[bool]['answer']
        
        # if there is D_index for the question append D_index to df
        if D_index.empty==False:
            df.loc[i,'D_index']=D_index.tolist()[0]
        
        # if there is D for the question append to D to df       
        if D.empty==False:
            df.loc[i,'D']=D.tolist()[0]
            
        ##########    for df_keys    ##############
        df['key_index']=df_keys['key_index']
        df['keys']=df_keys['key']
        
    return(df)


###################################
# Function to convert keys to answer_index
###################################
def KeysToAnswerIndex(keys, answer_index):
    """
    KeysToAnswerIndex function convert alphabet keys to matching answer index.
    :keys: list of keys. example [A,B,C]
    :answer_index: nested list of answer index. example [[1,2,3,4],[5,6,7,8],[9,10,11,12]]
    the example returns [1,6,11] 
    """
    # Construct dict to convert ABCD to 0123
    abcd_to_0123 = {'A':0, 'B':1, 'C':2, 'D':3}
    # convert ABCD keys to answer_index
    key_answer_index = []
    for i, key in enumerate(keys):
        key_n = abcd_to_0123[key]    # convert ABCD keys to 1234 keys
        key_index = answer_index[i][key_n]    # convert 1234 keys to keys
        key_answer_index.append(key_index)
    return(key_answer_index)


###################################
# Function to convert keys to answer_index
###################################
def AnswerIndexToKeys(key_answer_index, answer_index):
    """
    AnswerIndexToKeys function convert key answer to matching alphabet keys.
    :key_answer_index: list of key answer index. example [1,6,11]
    :answer_index: nested list of answer index. example [[1,2,3,4],[5,6,7,8],[9,10,11,12]]
    the example returns [A,B,C] 
    """
    # Construct dict to convert 0123 to ABCD
    num_to_ABCD = {0:'A', 1:'B', 2:'C', 3:'D'}
    # convert answer_index to keys
    newkeys = []
    for i, index in enumerate(answer_index):
        n_newkey = index.index(key_answer_index[i])
        newkey = num_to_ABCD[n_newkey]
        newkeys.append(newkey)
    return(newkeys)


###################################
# Function to update paragraphs
###################################
def UpdateAnswers(col='A_index', newstring='A. '):
    """
    UpdateAnswers function
    1. get the new_answers_index from [paragraphs] and replace it to [new_paragraphs].
    2. replace A. B. C. D. to correct order.
    :col: string of column name. Ex. 'A_index'
    :newstring: string of new start string Ex. 'A. '
    """
    # Get list of original and new answer line
    answer_index = df_to_shuffle[col].tolist()
    new_answers_index = df_shuffled[col].tolist()
    
    # Replace answer_index to new_answers_index
    for i, n in enumerate(new_answers_index):
        new_paragraphs[answer_index[i]] = paragraphs[n]
    
    ### Replace A. B. C. D. to A. B. C. D.
    for n in answer_index:
        
        # get the original start string. Ex. 'B. '
        p = new_paragraphs[n].find('w:t')
        string = UnicodeToString(p)
        # print((unicodedata.normalize('NFKD', new_paragraphs[n].get_text())))
        
        # create new start string. Ex. 'A. ' 
        string_new = newstring + string[3:]
        
        # replace the original start string with new start string
        p.string = string_new
        # print(unicodedata.normalize('NFKD', new_paragraphs[n].get_text()))
  
    print('Successfully replaced paragraphs with',col,'and replaced to', newstring)


###################################
# Function to copy file without document.xml
###################################
def CopyDocx(filein, fileout):
    """
    CopyDocx creates a copy of the docx file without 'word/document.xml'    
    :filein: string, original docx file name
    :fileout: string, new docx file name
    """
    with ZipFile(filein, 'r') as zipin:
        with ZipFile(fileout, 'w') as zipout:
            zipout.comment = zipin.comment # preserve the comment
            for item in zipin.infolist():
                if item.filename != 'word/document.xml':
                    zipout.writestr(item, zipin.read(item.filename))
    print(fileout,'is created')


###################################
# Function to make a new dataframe including keys
###################################
def KeyTable(df, df_shuffled):
    """
    KeyTable create a dataframe with question number, original key, and new key
    :df: dataframe containing original keys colume 'keys'
    :df_shuffled: dataframe containing new keys colume 'keys'
    """    
    # create list of question number in the docx
    n_questions = list(range(1,len(df)))
    # create dataframe containing question number column
    df_keytable = pd.DataFrame({'Q': n_questions})
    # add clumn containing original keys
    df_keytable[filein+' keys'] = df['keys']
    # add clumn containing new keys and add original key values
    df_keytable[fileout+' keys'] = df['keys']  
    # replace original key values to new key value for new key column
    for i, row in df_shuffled.iterrows():
        # print(i)
        # print(df_new_keys.loc[i,fileout+' keys'])
        # print(row['keys'])
        # print()
        df_keytable.loc[i,fileout+' keys'] = row['keys']
    
    return(df_keytable)


###################################
## Open and read docx
###################################
# filename
print('type filename with file extension. (Ex. Form A.docx)')
filein = input()
fileout = filein[:-5] + '_new' + filein[-5:]

#open docx with zipfile
with ZipFile(filein, 'r') as zip:
    doc = zip.read('word/document.xml')
    
# parse xml 
soup = BeautifulSoup(doc, 'xml')

# find body
body = soup.find('body')

# create Paragraph and save it to para
para = Paragraph(body)

# the last line number 
last_para_index = para.last()


###################################
## Extract Questions
###################################
# run get_questions method to extract Questions from para and save to dataframe
df_q = para.get_questions(search='\d+\. ', size=4)

# run AppendIndex function to append the last_para_index to the dataframe.
# it is needed for combining question and answer.
df_q = AppendIndex(df_q, last_para_index)

# print(df)


###################################
## Extract Answers
###################################
# run get_questions method to extract Questions from para and save to dataframe
df_a = para.get_answers(search='A\. ', size=4)
df_b = para.get_answers(search='B\. ', size=4)
df_c = para.get_answers(search='C\. ', size=4)
df_d = para.get_answers(search='D\. ', size=4)

# print(df_a)
# print(df_b)
# print(df_c)
# print(df_d)


###################################
## Extract Keys
###################################
# run get_keyss method to extract Keys from para and save to dataframe
searchstring = 'FORM A KEY'
df_keys = para.get_keys(search=searchstring, size=10)


###################################
## Combine Questions, Answers, Keys
###################################
# dataframe including questions, answers, and keys
df = JoinQandA(df_q, df_a, df_b, df_c, df_d, df_keys)

# dataframe to shuffle
df_to_shuffle= df.dropna()


###################################
## Generate Answers Index
###################################
col = ['A_index', 'B_index', 'C_index','D_index']
answer_index = df_to_shuffle[col].values.tolist()


###################################
## Generate Keys Answer Index
###################################
# get keys column from df_to_shuffle
keys = list(df_to_shuffle['keys'])

# run KeysToIndex function to convert keys to answer_index
key_answer_index = KeysToAnswerIndex(keys,answer_index)


###################################
## Shuffle Answers Index
###################################
for index in answer_index:
     random.shuffle(index)
#print(answer_index)


###################################
## Assign new keys
###################################
# run AnswerIndexToKeys function to convert key_answer_index to newkeys
newkeys = AnswerIndexToKeys(key_answer_index, answer_index)


###################################
## Construct new dataframe df_shuffled
###################################
# copy original data to new dataframe
df_shuffled = df_to_shuffle.copy()

# we only need index columns
df_shuffled.drop(['A','B','C','D'], axis=1, inplace=True)


###################################
## Assign new answers and keys
###################################
# get list of A_index, B_index, C_index, and D_index from the shuffled answer_index 
A_index=[]
B_index=[]
C_index=[]
D_index=[]

for index in answer_index:
    A_index.append(index[0])
    B_index.append(index[1])
    C_index.append(index[2])
    D_index.append(index[3])
    
# Assign the new line numbers of the answers to the columns
df_shuffled['A_index'] = A_index
df_shuffled['B_index'] = B_index
df_shuffled['C_index'] = C_index
df_shuffled['D_index'] = D_index

# Assign the new keys the columns
df_shuffled['keys'] = newkeys


###################################
## Construct new_paragraphs with new answers
###################################
# list of paragraph from docx
paragraphs = para.list
new_paragraphs = paragraphs.copy()

# run UpdateAnswers function to put paragraphs in correct name and order
UpdateAnswers(col='A_index', newstring='A. ')
UpdateAnswers(col='B_index', newstring='B. ')
UpdateAnswers(col='C_index', newstring='C. ')
UpdateAnswers(col='D_index', newstring='D. ')


###################################
## Construct New XML
###################################
# find a position of head from soup
n_head = str(soup).find('<w:body>')+8

# get string of xml head 
head = str(soup)[:n_head]

# find a position of tail from soup
n_tail = str(soup).find('</w:body>')
tail = str(soup)[n_tail:]

# slice new_paragraphs until right before keys section
questions_section_list = new_paragraphs[:last_para_index]
key_section_list = new_paragraphs[last_para_index:]

# join the list and convert to str
questions_section = ''.join(str(l) for l in questions_section_list)

# remove any text from key_section_list
for p in key_section_list:
    text = p.find_all('w:t')
    for tag in text:
        tag.clear()
    numlist = p.find_all('w:numPr')
    for tag in numlist:
        tag.clear()

# convert key_section_list to str
key_section = ''.join(str(l) for l in key_section_list)

# create soup_new by joining head, questions_section, and tail
soup_new =  head + questions_section + key_section + tail


###################################
## Save the new xml to docx
###################################
#run CopyDocx function to copy original docs file without 'word/document.xml'
CopyDocx(filein, fileout)

#add "word/document.xml' containing soup_new
with ZipFile(fileout, 'a') as newzip:
    newzip.writestr('word/document.xml', soup_new)
    print('word/document.xml', 'is added to', fileout)


###################################
## Save the keys to csv
###################################
# run KeyTable function to create key table
df_keytable = KeyTable(df, df_shuffled)

# save key table to csv
keyfilename = fileout[:-5] + '_key.csv'
df_keytable.to_csv(keyfilename, index=False)