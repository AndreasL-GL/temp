import os
class Config():
    SQL_CONNECTION =""
    SECRET_KEY = "SAOMDK#)1ek 0=#"
    API_KEYS = ['ABCDEFG','Nisse','Matsnyckel']
    ACCEPT_CONNECTIONS_FROM = []

class ImageConfig():
    image_size_after_resize = (300,300)

class SharepointConfig():
    with open(os.path.join(os.path.join(os.path.dirname(__file__),'config'),'sharepoint_egenkontroller_remove_list.txt')) as f:
        remove_list = [x.replace('\n','') for x in f.readlines() if x!='\n']

if __name__ == '__main__':
    print(SharepointConfig.remove_list)