'''
@Date: 2019-10-16 15:05:23
@LastEditors: cany
@Author: xmpan
@LastEditTime: 2019-10-22 17:25:27
'''
import xlwings as xw
import pandas as pd
import os


def convertCsv(filecsv):
    '''
    Read analysis CSV results,and convert to pandas dataframe write into Excel,also output full-length sequence

    Example:
    convertCsv(r'D:/1A/Synthetic_library_design/GermlineGrouping20190919/test/HV1-69.fullLength.17aa.csv')

    output: tab2w(pandas dataframe),seq('QVQLVQSGAEEVKKPGSSVKVSCKASGGTFSSYAISWVRQAPGQGLEWMGGIIPIFGTANYAQKFQGRVTITADESTSTAYMELSSLRSEDTAVYYCARDGGYSGGYYYYYFDVWGQGTLVTVSS')
    '''

    aaColTemp = [
        'A', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'K', 'L', 'M', 'N', 'P', 'Q',
        'R', 'S', 'T', 'V', 'W', 'Y'
    ]

    myt = pd.read_csv(filecsv, index_col=0).T

    mytable = myt.replace(0, '')
    aaDict = {'aa': aaColTemp, 'x': aaColTemp}
    df = pd.DataFrame(aaDict).set_index('aa')
    mydata = df.join(mytable).drop('x', axis=1)

    tab2w = mydata[mydata != 1].dropna(axis=1, how='any')

    # each column most popular AA ---> seqtemplate ,as excel template # Wild Type ORF Sequence - DNA
    seqlist = myt.idxmax().tolist()
    seq = ''.join(seqlist)

    return tab2w, seq


def writeIntoExcel():
    '''
    Write info to the defined location. Xlwings support record lists directly.    
    '''

    tab2w, seq = convertCsv(csvFile)

    wb = xw.Book(fileout)
    sht = wb.sheets['CVL Input']

    sht.range('D13').value = '#'
    sht.range('D39').value = seq

    #ORF AA POSITION
    sht.range('H10').value = tab2w.columns.tolist()

    #https://codeday.me/bug/20180823/227073.html，record lists
    sht.range('H13').value = tab2w.values.tolist()

    #save excel
    print('xls save as ' + fileout)
    wb.save(fileout)
    wb.close()


if __name__ == "__main__":
    import sys, subprocess
    csvFile = sys.argv[1]
    temp = r'D:\1A\Synthetic_library_design\GermlineGrouping20190919\test\SubmissionForm_Library_CVL_17Jul2019_v3.6.xlsx'

    outname = os.path.basename(csvFile).split('.csv')[0]
    fileout = f'{outname}.xlsx'
    subprocess.getoutput(f'cp {temp} {fileout}')

    writeIntoExcel()