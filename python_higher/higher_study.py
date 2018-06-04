# -*- coding:utf-8 -*-

"""
create on 2018-05-21
@author linwei

study the python higher attribute

"""
#if you do not to know how many paras to be passed on， then using *args(可变的参数列表) or **kwargs（可变的键值对）
def adapt_para(para, *args):
    print 'the first value is:', para
    print type(args)
    for i in args:
        print 'the other value is :', i

def adapt_para1(**kwargs):
    for key, value in kwargs.items():
        print 'the param is {0} and {1}:'.format(key, value)

adapt_para(123, 'abc', 'dfdf', 'dfd')

adapt_para1(name='abc', age='18')

#learn to use the generator   yield
def generator_use():
    for i in range(100):
        yield i
# for item in generator_use():
#     print item
gen = generator_use()
print next(gen)
print next(gen)

str1 = 'abcdef'
str_generator = iter(str1)
print next(str_generator)


#learn to using map,  list map to def
items = [1, 2, 3, 4, 5, 6]
print list(map(lambda x:x*x, items))

def add(x):
    return x+x
def mul(x):
    return x*x
for i in range(5):
    print list(map(lambda x:x(i), [add, mul]))





