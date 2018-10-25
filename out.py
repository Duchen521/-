# -*- coding: UTF-8 -*-
import wordproduce2
ww = u'样好大家好这'
for i in range(0,6):
        pathway = 'E:\\sampleProduce\\' + str(i) + ".docx"
        topath = 'E:\\sampleProduce\\' + str(i+1) + ".docx"
        my_word = wordproduce2.WordWrap(pathway)
        my_word.textReplace(ww[i],ww[i+1])
        my_word.saveAs(topath)
        my_word.quit()
        del my_word
        print i+1
