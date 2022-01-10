# -*- coding: utf-8 -*-
"""
Created on Wed Aug 15 13:54:58 2018

@author: HP
"""

# -*- coding: utf-8 -*-
"""
Created on Mon Aug 13 13:27:01 2018

@author: HP
"""

#有道词典2018-8-13 13:27:10
#实现单词查询功能

import re
import urllib.request


 #返回html的字符串
def getHtml (url): 
    req = urllib.request.Request(url=url,headers={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) '  
                        'Chrome/51.0.2704.63 Safari/537.36'}  )
    page = urllib.request.urlopen(req)
    htmlstr = page.read().decode(encoding='utf-8',errors='strict')
    return htmlstr



def get_shiyi(word):
    #r'<ul> {0,12}\n {0,12}( {0,12}<li>\w+\.[u4e00-u9fa5]*?</li> {0,12}\n {0,12})+ {0,12}</ul>'
    shiyi_pattern = re.compile(r'<li>\w.*</li>?')
    url = ("http://dict.youdao.com/w/%s")%word
    html = getHtml(url)
    #html = html[:5000]
    a = html.find('<div class="trans-container">')
    html = html[a:a+1500]
    t = re.findall(shiyi_pattern,html)
    result = []
    for each in t:
        result.append(each[4:-5])
    if result == []:
        result = ["查无此词，请检查拼写"]
    else:
        #result = list(set(result))
        pass
    return result#type==list
#print(html)


if __name__ == "__main__":
    word = ""
    t = 1
    while 1>0:
        if word == "" or t>1:
            word = input("请输入要查找的单词：")
        shiyi = get_shiyi(word)
        print (word+":")
        for each in shiyi:
            print (each)
        t += 1

















