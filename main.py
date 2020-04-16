import docx
import re
from os import listdir
from os.path import isfile, join
import csv

result_file = "result.csv"
filesPath = './documents/'


words = []

def para_to_text(p):
    rs = p._element.xpath('.//w:t')
    return u" ".join([r.text for r in rs])


def extract_words(my_doc):
    for para in my_doc.paragraphs:
        texts = para_to_text(para)
        p = re.compile(r'[a-zA-Z\']+')
        matches = p.findall(texts)
        for m in matches:
            words.append(str.lower(m))


def extract_file(f):
    if f.startswith("CET") and f.endswith("docx"):
        path = filesPath + f
        doc = docx.Document(path)
        extract_words(doc)


def list_to_freq_dict(word_list):
    word_freq = [word_list.count(p) for p in word_list]
    print("listed")
    return dict(list(zip(word_list, word_freq)))


def sort_freq_dict(freq_dict):
    aux = [(freq_dict[key], key) for key in freq_dict]
    aux.sort()
    aux.reverse()
    print("sorted")
    return aux


files = [f for f in listdir(filesPath) if isfile(join(filesPath, f))]

for file in files:
    extract_file(file)

print(words.__len__())

words_freq = list_to_freq_dict(words)
sorted_word_freq = sort_freq_dict(words_freq)

with open(result_file, "a") as csvfile:
    writer = csv.writer(csvfile)
    for w in sorted_word_freq:
        writer.writerow(list(w))

    print("done")
    csvfile.close()
