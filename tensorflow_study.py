# -*- coding:utf-8 -*-

"""
create on 2018-05-23
@author linwei

study the tensorflow

"""
import tensorflow as tf
import os

os.environ['TF_CPP_MIN_LOG_LEVEL'] = '2'
def create_NN1():
    x = tf.constant([[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]])
    w = tf.constant([[5.0, 7.0], [6.0, 8.0], [7.0, 9.0]])
    y = tf.matmul(x, w)

    x1 = tf.constant([[1.0, 2.0]])
    w1 = tf.constant([[5.0], [6.0]])
    y1 = x1 + w1

    print x
    print y
    print y1

    with tf.Session() as sess:
        print sess.run(y)
        print sess.run(y1)

# x:1,2  w1:2,3  a1: 1,3  w2:3,1  y
def create_NN2():
    #定义输入和参数
    x = tf.placeholder(tf.float32,shape=(None,2))
    w1 = tf.Variable(tf.random_normal([2,3],stddev=1,seed=1))
    w2 = tf.Variable(tf.random_normal([3,1],stddev=1,seed=1))

    #定义前向传播过程
    a1 = tf.matmul(x, w1)
    y = tf.matmul(a1, w2)

    with tf.Session() as sess:
        init_op1 = tf.global_variables_initializer()
        sess.run(init_op1)
        print sess.run(y, feed_dict={x:[[0.7,0.5],[0.4,0.5],[0.3,0.2]]})

create_NN2()

