"""
create on 2018-4-15
@author linwei

"""

#字典常用方法
def dict():

    my_dict = {"a":1, "b":2, "c":3}
    print my_dict.has_key("b")

    print my_dict.items()
    print my_dict.keys()
    print my_dict.values()

#统计单词中字母出现的个数。
def str_count():
    str = "aaabbccccc"
    my_dict = {"a":0, "b":0, "c":0}
    for i in str:
        if my_dict.has_key(i):
            my_dict[i] += 1
    print my_dict

#统计字母出现的个数，使用更好的方法
def str_count_pang(content):
    count = {}
    for character in content:
        if character not in count:
            character_count = content.count(character)
            count.update({character: character_count})
            print count
    # print count

#统计一句话中出现最多的前两个单词
def count_words():
    str = "avc  abc  sdfd  aa  vv  xx  aa  xx   aa  aa  vv  bb "
    str_list = str.split()
    count_dict = {}
    for word in str_list:
        if word not in count_dict:
            counts = str_list.count(word)
            count_dict.update({word:counts})

    list_counts = []
    for key, value in count_dict.items():
        list_counts.append((value, key))
    list_counts.sort(reverse= True)
    for key, value in list_counts[:2]:
        print value, key

#经典的反转字典的使用。
def reverse_dict():
    my_dict = {"linwei": 123, "liulei": 123, "lilei": 234, "hanmeimei": 234, "david": 456}
    other_dict = {}
    for name, number in my_dict.items():
        if number in other_dict:
            other_dict[number].append(name)
        else:
            other_dict[number] = [name]
    print other_dict

str_count()
str_count_pang("abcabcefefefekj")
count_words()
reverse_dict()