import os, sys
import shutil
import docx
import re
from sklearn.feature_extraction.text import CountVectorizer
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import comtypes.client
#import pdfminer.six


class Converter(object):
    def __init__(self,):
        self.compiled_re = re.compile('(?<=[.]).*')
        self.doc_or_docx2pdf_files = []

    def run(self, fname, hw_dirname):
        print('{}'.format(fname))
        chg_fname = self.compiled_re.sub('txt', fname)
        fext_name = re.search('(?<=[.]).*', fname).group(0)

        if ('doc' == fext_name) or ('docx' == fext_name):
            # convert .docx to .pdf
            chg2pdf_fname = self.docx2pdf(fname, fext_name)
            self.pdf2txt(hw_dirname, chg_fname, chg2pdf_fname)
        elif 'pdf' == fext_name:
            self.pdf2txt(hw_dirname, chg_fname, fname)

    def docx2pdf(self, filename, fext_name):
        wdFormatPDF = 17
        in_file = os.getcwd() + '/{}'.format(filename)
        out_file = os.getcwd() + '/{}'.format(filename.replace('.{}'.format(fext_name), '.pdf'))
        chg2pdf_fname = filename.replace('.{}'.format(fext_name), '.pdf')

        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()
        self.doc_or_docx2pdf_files.append(chg2pdf_fname)
        return chg2pdf_fname

    def pdf2txt(self, hw_dirname, chg_fname, pdf_fname):
        output_filepath = './{}/{}'.format(hw_dirname, chg_fname)

        pdf2txt_path = re.search('.*(?=\\\python)', sys.executable).group(0)
        pdf2txt_path = pdf2txt_path.replace('\\', '/')
        if 'Scripts' not in pdf2txt_path:
            pdf2txt_path = pdf2txt_path + '/Scripts'

        if os.path.isfile(pdf2txt_path + '/pdf2txt.py'):
            os.system('{} {}/pdf2txt.py {} -o {}'.format(
                sys.executable,
                pdf2txt_path,
                pdf_fname,
                output_filepath))
        else:
            raise ImportError('Please do "pip install pdfminer.six" -> pdf2txt converting library')

    def del_doc_or_docx2pdf_fils(self):
        for fn in self.doc_or_docx2pdf_files:
            os.unlink(os.getcwd() + '/{}'.format(fn))

class CopiedDocumentChecker(object):
    """
    It accepts file extensions of .docx, .pdf, .hwp only.
    .hwp parsing is prone to errors.
    """

    def __init__(self, dirpath):
        self.allowed_ext_names = ['doc', 'docx', 'pdf']
        self.dirpath = dirpath
        self.convter = Converter()
        self.hw_dirname = 'converted_hws'

    def ch_cwd(self):
        os.chdir(self.dirpath)

    def check_fext(self):
        fext_names = os.listdir('./')
        self.sot_fext_names = []
        for fext_name in fext_names:

            try:
                fext_name_ext = re.search('(?<=[.]).*', fext_name).group(0)
            except AttributeError:  # it occurs when it's parsing a folder
                pass

            if fext_name_ext not in self.allowed_ext_names:
                raise TypeError('Extension type is violated -> {}'.format(fext_name))
            else:
                if ('.' in fext_name) and ('~$' not in fext_name[:2]):
                    self.sot_fext_names.append(fext_name)

    def cvt_files2txt(self):
        if not os.path.isdir('./' + self.hw_dirname):
            os.mkdir('./' + self.hw_dirname)
        else:
            shutil.rmtree('./' + self.hw_dirname)
            os.mkdir('./' + self.hw_dirname)

        for idx, sfn in enumerate(self.sot_fext_names):
            print("converting the file to .txt file ... | {}/{}".format(idx + 1, len(self.sot_fext_names)), end=' | ')
            self.convter.run(sfn, self.hw_dirname)

    def get_docname8content_dict(self):
        covt_fnames = os.listdir('./' + self.hw_dirname)

        self.docname8content_dict = {}
        for cvtfn in covt_fnames:

            content = []
            try:
                with open('./{}/{}'.format(self.hw_dirname, cvtfn), 'rb') as f:
                    for line in f:
                        content.append(line.decode())
            except:
                try:
                    with open('{}/{}'.format(self.hw_dirname, cvtfn), 'r') as f:
                        for line in f:
                            content.append(line)
                except:
                    pass

            content = ''.join(content)
            self.docname8content_dict[cvtfn] = content

    def get_word_count_matrix(self, ngram_range=(2,2), min_df=1):
        vect = CountVectorizer(ngram_range=ngram_range, min_df=min_df)
        self.word_count_mat = vect.fit_transform( list(self.docname8content_dict.values()) ).toarray()

    def get_ecl_dist_matrix(self, penalty_to_similarity=10):
        ecl_dist_matrix = np.zeros((self.word_count_mat.shape[0], self.word_count_mat.shape[0]))

        for i in range(ecl_dist_matrix.shape[1]):
            print('Calculating the similarities | {:0.1f}%'.format((i + 1) / ecl_dist_matrix.shape[1] * 100))
            for j in range(ecl_dist_matrix.shape[0]):
                target_word_count_mat = self.word_count_mat[i]
                compared_word_count_mat = self.word_count_mat[j]

                # 둘다 한 word feature 에서 값을 가지고있는경우에는,
                # additional penalty 를 가한다. 왜냐하면 이런 경우가 copying 일 가능성이 크기때문.
                target_bool = target_word_count_mat >= 1
                compared_bool = compared_word_count_mat >= 1
                cond_satisf = []
                cond_count_gap_plus_one = []
                for tb, cb, t, c in zip(target_bool, compared_bool, target_word_count_mat, compared_word_count_mat):
                    if tb == True and cb == True:
                        cond_satisf.append(1)
                        cond_count_gap_plus_one.append(np.abs(t - c) + 1)
                    else:
                        cond_satisf.append(0)
                        cond_count_gap_plus_one.append(1)

                norm_ecl_dist = np.sqrt(np.sum((target_word_count_mat - compared_word_count_mat) ** 2))

                if not i == j:
                    norm_ecl_dist += np.sum((-penalty_to_similarity * np.array(cond_satisf)) / np.array(
                        cond_count_gap_plus_one))  # apply the penalty

                ecl_dist_matrix[i, j] = norm_ecl_dist

        self.df_ecl_dist_matrix = pd.DataFrame(ecl_dist_matrix,
                                               columns=list(self.docname8content_dict.keys()),
                                               index=list(self.docname8content_dict.keys()))
        return self.df_ecl_dist_matrix

    def catch_cheaters(self, n_top_likely=15):
        arr = self.df_ecl_dist_matrix.values
        df_shape = self.df_ecl_dist_matrix.values.shape

        sot_arr_dict = {}
        for i in range(df_shape[1]):
            for j in range(df_shape[0]):

                if i <= j:
                    continue
                else:
                    sot_arr_dict[(i, j)] = arr[i, j]

        sot2_arr_dict = sorted(sot_arr_dict.items(), key=lambda x: x[1])


        self.reset_result_dir()
        print("\n\n================ List of Cheating Suspects ================")
        for idx, val in enumerate(sot2_arr_dict[:n_top_likely]):
            coord = val[0]
            ecl_dist = val[1]
            chaeting_suspect1 = self.df_ecl_dist_matrix.columns[coord[0]]
            cheating_suspect2 = self.df_ecl_dist_matrix.index[coord[1]]
            top_ranking = idx + 1

            print("Cheating Suspect Top-{} | ecl_dist(=similarity): {:0.1f}".format(top_ranking, ecl_dist))
            print('{} <-> {}'.format(chaeting_suspect1, cheating_suspect2), end='\n\n')

            # save the cheating suspects' docs in the result folder.
            self.save_cheating_docs(chaeting_suspect1, cheating_suspect2, top_ranking)
        print("===========================================================")
        print("[NOTE] The likely cheating files are saved in the '{}/result' folder. \n"
              "Please check it and screw them haha.".format(self.dirpath))

    def plot_heatmap(self, slice_fname_for_plt=8, plot_=False, save=True):
        plt.figure(figsize=(15, 10))

        plt.rcParams['xtick.bottom'] = plt.rcParams['xtick.labelbottom'] = False
        plt.rcParams['xtick.top'] = plt.rcParams['xtick.labeltop'] = True

        plt.pcolor(self.df_ecl_dist_matrix[::-1], cmap='coolwarm')
        df_shape = self.df_ecl_dist_matrix.shape

        for i in range(df_shape[1]):
            for val_j, plot_j in zip(range(df_shape[0]), range(df_shape[0] - 1, 0 - 1, -1)):
                plt.text(x=i + 0.5, y=plot_j + 0.5, s='{:0.1f}'.format(self.df_ecl_dist_matrix.iloc[i, val_j]),
                         horizontalalignment='center', verticalalignment='center')

        xticks_ = list(map(lambda x: x[:slice_fname_for_plt], self.df_ecl_dist_matrix.index.tolist()))
        yticks_ = list(map(lambda x: x[:slice_fname_for_plt], self.df_ecl_dist_matrix.index.tolist()))

        plt.yticks(np.arange(0.5, len(self.df_ecl_dist_matrix.index), 1), yticks_[::-1])
        plt.xticks(np.arange(0.5, len(self.df_ecl_dist_matrix.columns), 1), xticks_, rotation=90)
        plt.xlabel('Normalized Euclidean Distance Matrix\n'
                   'The lower the euclidean distance, the higher the similarity btn the two documents\n'
                   'Cheating Chance = 1 / Euclidean Distance', size=15)

        _ = plt.show() if plot_ else None

        _ = plt.savefig('./result/euclidean_dist_matrix.png') if save else None

    def reset_result_dir(self):
        # clear the result folder
        if not os.path.isdir('./result'):
            os.mkdir('./result')
        else:
            shutil.rmtree('./result')
            os.mkdir('./result')

    def save_cheating_docs(self, ssp1, ssp2, top_rank):
        ssp1, ssp2 = re.sub('(?<=).txt', '', ssp1), re.sub('(?<=).txt', '', ssp2)

        # find the original ext. names
        for fn in self.sot_fext_names:
            if ssp1 in fn:
                ssp1_ori_fname = fn
            if ssp2 in fn:
                ssp2_ori_fname = fn

        # make a individual dir. for top-n
        os.mkdir('./result/Top-{}'.format(top_rank))

        # save in the result dir.
        shutil.copyfile(ssp1_ori_fname, './result/Top-{}/'.format(top_rank) + ssp1_ori_fname)
        shutil.copyfile(ssp2_ori_fname, './result/Top-{}/'.format(top_rank) + ssp2_ori_fname)

    def run(self, n_top_likely, penalty_to_similarity=10):
        self.ch_cwd()
        self.check_fext()
        self.cvt_files2txt()
        self.get_docname8content_dict()
        self.get_word_count_matrix()
        self.get_ecl_dist_matrix(penalty_to_similarity)
        self.convter.del_doc_or_docx2pdf_fils()

        self.catch_cheaters(n_top_likely)
        self.plot_heatmap()


if __name__ == '__main__':

    # directory of a folder that contains homework files.
    dirpath = './students_homeworks_example'

    # run
    checker = CopiedDocumentChecker(dirpath)
    checker.run(n_top_likely=15)
