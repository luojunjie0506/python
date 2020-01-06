import wordcloud
import  jieba
import  imageio
import itchat
'''
#登录微信
itchat.login()
tList = []
#获取好友列表
friends= itchat.get_friends(update=True)
#获取所有好友的个签存入tList列表中
for i in friends:
    #获取个签
    signature = i['Signature']
    if 'emoji' in signature:
        pass
    else:
        tList.append(signature)
text = ' '.join(tList)
'''
#打开并读取文本
text = open('H:\\P\\aa.txt',encoding='utf-8').read()
#文本分词，成列表
a = jieba.lcut(text,cut_all = True)
b = ' '.join(a)

#新建图片对象
mk = imageio.imread('c.jpg')
#新建词云对象
w = wordcloud.WordCloud(width=1000,
                        height=700,
                        background_color='white',#背景颜色
                        font_path='msyh.ttc',#字体
                        mask=mk,#背景图
                        scale=15,#清晰度
                        )
#传入文本
w.generate(b)
w.to_file('H:\\P\\ciyun.PNG')