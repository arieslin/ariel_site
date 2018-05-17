import multiprocessing
import time

def test(num):
    print 'test:', num


# if __name__ == '__main__':
#     print time.time()
#     for i in range(10):
#         process = multiprocessing.Process(target=test, args=(i,))
#         process.start()
#     print time.time()
# print 'aaa'
print time.time()
for i in range(10):
    test(i)
print time.time()