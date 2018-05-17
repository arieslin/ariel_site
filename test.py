

def xyz():
    x = 1
    y = 2
    z = 3
    x =  y
    y = x
    z = y

    print z,x ,y

def count():
    count = 0
    while count < 5:
        print count
        count += 1

def all(end):
    i = 1
    s = 0
    while i <= end:
        s = s + i
        i += 1
    print "the value of s:", s

def break_test():
    cnt = 0
    while cnt < 5:
        cnt += 1
        if cnt > 2:
            continue
        print "hello"

def for_use():
    s = 0
    for i in range(1000):
        s = s + i
    print s

def pi():
    pi = 0
    for i in range(1, 1000):
        pi += (-1.0)**(i + 1) / (2*i - 1)
    print pi*4

def collaz():
    n = 6
    while n != 1:
        if n%2 == 0 :
            n = n/2
        else:
            n = n*3 + 1
        print n
def count17():
    max = 10
    sum = 0
    extra = 0
    for num in range(1,max):
        if num % 2 and not num % 3:
            sum += num
            print num
        else:
            extra += 1
            #print extra
    print sum
    print extra
# pi()
# for_use()
# count()
# all(10)
# break_test()
# collaz()
count17()

