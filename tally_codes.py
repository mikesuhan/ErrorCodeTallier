import os
import re
import docx2txt


class Tallier:
    """Generates two csv files for the raw and normalized frequencies of substrings in a .docx file's text.
    
    Files should be named in a consistent pattern identifying the treatment and subject, because this will
    determine how the output is organized. For example:
    
    1.Jimmy.docx
    1.Sally.docx
    2.Jimmy.docx
    2.Sally.docx

    Arguments:
        
        folder: A folder of .docx files
        output_filename: The first part of the output file's filename. 
        
    Keyword Arguments:
        
        key_text: A string after which point error codes are not counted anymore. For example, maybe there is an
        answer key at the end of a text that shows the subject the meaning of the codes, but the codes in it
        do not represent errors made by the subject.
        
        delimiter: The string used to separate the subject/case from the treatment. e.g. "." in 1.Jimmy.docx

        code_file: The filepath to the file in which a csv of the error codes are stored. If None is used,
        the codes class variable will be used in its place.
        
        code_file_encoding: Character encoding of code_file.
        
        treatment_ind: An integer representing the place of the treatment label in the docx filenames.
        
        case_ind: An integer representing the index of the case/subject in the docx filenames.
    
    Example:
        
        Creates MyCSV_frequencies.csv and MyCSV_normalized.csv from docx files in the example_texts directory.
        
        >>> from tally_codes import Tallier
        >>> t = Tallier('example_texts', 'MyCSV')
        >>> t.process()
    """
    # default error codes used if no file is loaded
    codes = {'[AE]':'article_error',
             '[RO]':'run_on_sentence',
             '[WF]':'word_form_error',
             '[VT]':'verb_tense_error',
             '[PL]':'plural_error',
             '[SVA]':'3rd_person_error'}

    left_bound, right_bound = '[', ']'

    def __init__(self, folder, output_filename, key_text='Key for Error Types:', delimiter='.', code_file='codes.csv',
                 code_file_encoding='UTF-8', treatment_ind=0, case_ind=1):
        self.folder = folder
        self.key_text = key_text
        self.delimiter = delimiter
        self.treatment_ind = treatment_ind
        self.case_ind = case_ind
        self.files = [f for f in os.listdir(folder) if not f.startswith('~')]
        self.output_filename = output_filename

        if code_file:
            self.codes = self.read_codes(code_file, code_file_encoding)

    def read_codes(self, code_file, code_file_encoding):
        """Defines error codes based on csv data.
        
        Arguments:
            code_file: the file the error codes are stored in
            code_file_encoding: the encoding of code_file

        CSV format example:
        
        [AE], article_error
        [RO], run_on_sentence
        [WF], word_form_error
        """
        with open(code_file, encoding=code_file_encoding) as f:
            text = f.read().splitlines()

        # separates codes from labels
        codes = [line.split(',') for line in text if line.strip()]
        codes = [(item[0], ' '.join(item[1:]).strip()) for item in codes]
        codes = {code: label for code, label in codes}

        return codes


    def tally(self, rate=100):
        """Counts the error codes in each file, returning a string in csv format.
        
        Keyword arguments:
            rate: The rate at which error code counts will be normalized to using: n / word_count * rate
        """

        docs = []

        for file in self.files:
            print(file)
            text = self.docx_text(self.folder + '/' + file)
            fd = {e: text.count(e) for e in self.codes}

            word_count = self.docx_word_count(text)
            file = file.split(self.delimiter)
            treatment_label = file[self.treatment_ind]
            case_label = file[self.case_ind]
            # treatment label, case label, frequency distribution
            docs.append([file[0], file[1], fd, word_count])

        if not self.files:
            raise Exception('No files have been loaded.')

        treatments = sorted(set(int(tr) for tr, _, __, ___ in docs))
        cases = sorted(set(c for _,c,__, ___ in docs))

        # table delimiter
        tdl = ', '

        # makes column header for each treatment
        tr_header = ''

        for tr in treatments:
            for c in self.codes:
                tr_header += tdl + self.codes[c] + '_' + str(tr)
            tr_header += tdl + 'word_count_' + str(tr)


        header = 'subjects' + tr_header + '\n'
        table = header

        # makes row for each case
        for case in cases:
            row = case
            for treatment in treatments:

                item = [(tr,c,frd,wc) for tr, c, frd, wc in docs if int(tr) == treatment and c == case]
                if item:
                    for code in self.codes:
                        tr,c,frd,wc = item[0]
                        # Adds error counts
                        freq = frd.get(code, 0)
                        if rate:
                            row += tdl + str(round(freq/word_count*rate, 2))
                        else:
                            row += tdl + str(freq)

                    # Adds word count
                    row += tdl + str(item[0][3])

                else:
                    # Adds blank spaces if there is no treatment for a case. + 1 because of word_count
                    row += tdl * (len(self.codes) + 1)
            row += '\n'
            table += row
        return table


    def docx_text(self, filepath):
        '''Tokenizes the text in a .docx file (without punctuation)'''

        text = docx2txt.process(filepath)
        try:
            cut_off = text.index(self.key_text)
            return text[:cut_off]
        except ValueError:
            print('No answer key present in', filepath)
            return text


    def docx_word_count(self, text):
        """Returns the word count of a docx file."""
        text = re.sub('[^A-Za-z\s0-9-\']', '', text)
        return len(text.split())

    def process(self):
        """Runs everything and saves it as a csv."""
        with open(self.output_filename + '_frequencies.csv', 'w', encoding='utf8') as f:
            f.write(self.tally(False))
        with open(self.output_filename + '_normalized.csv', 'w', encoding='utf8') as f:
            f.write(self.tally())
                
    



    
    
