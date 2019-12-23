import os
path = "C:\Yogesh\Repository\cdp\src\main\java\com\tmp\cdp"
myCmd = "dir "+ path
os.system(myCmd)

myCmd = os.popen(myCmd).read()

myCmd = myCmd.split('\n')

for i in myCmd:
    str = i.split('<DIR>')
    if len(str) > 1:
         str = str[1].strip()
         if str == "." or str ==".." :
            continue
         print("directory " + str)
         os.chdir(path)
         os.system("cd "+str) # here is working directory path needs to add
         res = os.popen("dir "+str).read()
         print(res)
         os.system("cd ..")


#print(type(myCmd))
