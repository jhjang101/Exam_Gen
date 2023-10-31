# import
import util

def main():
    file, forms = util.file_in() # user prompt. returns filename and list of forms
    soup = util.read_docx(file) # parse xml
    df, df_key = util.read_questions(soup) # read questions and keys

    # shuffle and save 
    for form in forms:
        df_to_shuffle, df_shuffled = util.shuffle(df) # shuffle
        paragraphs = util.to_paragraphs(soup,df_to_shuffle,df_shuffled,form) # reorder paragraphs
        xml = util.to_xml(soup, paragraphs) # rebuild xml
        df_key = util.update_keytable(df_key, df_shuffled,form) # update key table
        util.to_file(xml,df_key,file,form) # save
    
    print(df_key)
        


if __name__ == "__main__":
    main()