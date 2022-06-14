from email import policy
from email.parser import BytesParser
from glob import glob
import os

# 해당 VIN No.와 동일한 EMAIL파일의 경로를 찾는다.

currentAbsPath = os.path.dirname(os.path.realpath(__file__))
file_list = list(glob(f"{currentAbsPath}/upload/EMAIL/*.eml"))

for file in file_list:
    with open(file, 'rb') as fp:
        msg = BytesParser(policy=policy.default).parse(fp)
        txt = msg.get_body(preferencelist=('plain')).get_content()
        if txt.find("W1K3F4EB4MJ319419") > -1:
            print(type(file))


# for file in emlFiles:
#     with open(file, 'rb') as fp:
#         print(fp.read())
#         input('pause:')
