

file1 = open("/Users/arieslin/python/test/aaa.txt","r")
filelist = file1.readlines()
for x in filelist:
    print x.rstrip('\n')

file1.close()


list1 = ['python1', 'python2','python3']
file2 = open('/Users/arieslin/python/test/abc.txt', 'w')

for i in list1:
    file2.writelines(i)
    file2.writelines('\n')
file2.close()

def test1(x,y,z):
    return x , y, z
x = test1(3,4,6)
print x

