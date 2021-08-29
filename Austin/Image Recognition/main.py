import tensorflow as tf
from tensorflow.examples.tutorials.mnist import input_data
import numpy as np
from PIL import Image

#https://www.digitalocean.com/community/tutorials/how-to-build-a-neural-network-to-recognize-handwritten-digits-with-tensorflow

mnist = input_data.read_data_sets("MNIST_data/", one_hot=True)  # y labels are oh-encoded

# number of example data points
n_train = mnist.train.num_examples  # 55,000
n_validation = mnist.validation.num_examples  # 5000
n_test = mnist.test.num_examples  # 10,000

# Define Layers of neural network
n_input = 784  # input layer (28x28 pixels)
n_hidden1 = 512  # 1st hidden layer
n_hidden2 = 256  # 2nd hidden layer
n_hidden3 = 128  # 3rd hidden layer
n_output = 10  # output layer (0-9 digits)

# Define Hypervariables
learning_rate = 1e-4
n_iterations = 1000
batch_size = 128
dropout = 0.5

# define 3 tensors as placeholders
X = tf.placeholder("float", [None, n_input])  # feed in a None (unknown) amount of n_input (784) pixel images
Y = tf.placeholder("float", [None, n_output]) # feed out unknown amount of n_output (10) possible outputs
keep_prob = tf.placeholder(tf.float32) # we inititailze keep_prob as a placeholder so we can use it with a dropout rate
                                        # of 0.5 now, and later 1.0 when we test

# weights are randomly selected from a truncated normal distrobution (better accuracy than if we unradomly set them)
weights = {
    'w1': tf.Variable(tf.truncated_normal([n_input, n_hidden1], stddev=0.1)),
    'w2': tf.Variable(tf.truncated_normal([n_hidden1, n_hidden2], stddev=0.1)),
    'w3': tf.Variable(tf.truncated_normal([n_hidden2, n_hidden3], stddev=0.1)),
    'out': tf.Variable(tf.truncated_normal([n_hidden3, n_output], stddev=0.1)),
}

# for biases, we use a constant rather than a random number, to make sure that tensors actually activate during the initial
# training iterations
biases = {
    'b1': tf.Variable(tf.constant(0.1, shape=[n_hidden1])),
    'b2': tf.Variable(tf.constant(0.1, shape=[n_hidden2])),
    'b3': tf.Variable(tf.constant(0.1, shape=[n_hidden3])),
    'out': tf.Variable(tf.constant(0.1, shape=[n_output]))
}

#Each hidden layer will execute matrix multiplication on the previous layer’s outputs and the current layer’s weights,
# and add the bias to these values. At the last hidden layer, we will apply a dropout operation using our keep_prob
# value of 0.5.
layer_1 = tf.add(tf.matmul(X, weights['w1']), biases['b1'])
layer_2 = tf.add(tf.matmul(layer_1, weights['w2']), biases['b2'])
layer_3 = tf.add(tf.matmul(layer_2, weights['w3']), biases['b3'])
layer_drop = tf.nn.dropout(layer_3, keep_prob)
output_layer = tf.matmul(layer_3, weights['out']) + biases['out']


# define the loss function to optimize
cross_entropy = tf.reduce_mean(  #cross_entropy is a popular tensor-flow loss function
    tf.nn.softmax_cross_entropy_with_logits(
        labels=Y, logits=output_layer
        ))
train_step = tf.train.AdamOptimizer(1e-4).minimize(cross_entropy) #Adam optimizer is a type of gradient-descent optimizer
                                                                    # that can minimize the loss function

# define what is correct vs incorrect, and what is accuracy
correct_pred = tf.equal(tf.argmax(output_layer, 1), tf.argmax(Y, 1)) # see if training guess is equal to stored value (1 or 0)
accuracy = tf.reduce_mean(tf.cast(correct_pred, tf.float32))  # get % accurate by averaging the booleans

#initialize training session
init = tf.global_variables_initializer()
sess = tf.Session()
sess.run(init)



# train on mini batches
# We use mini-batches of images rather than feeding them through individually to speed up the training process and
# allow the network to see a number of different examples before updating the parameters.
for i in range(n_iterations):
    batch_x, batch_y = mnist.train.next_batch(batch_size)
    sess.run(train_step, feed_dict={
        X: batch_x, Y: batch_y, keep_prob: dropout
        })

    # print loss and accuracy (per minibatch)
    if i % 100 == 0:
        minibatch_loss, minibatch_accuracy = sess.run(
            [cross_entropy, accuracy],
            feed_dict={X: batch_x, Y: batch_y, keep_prob: 1.0}
            )
        print(
            "Iteration",
            str(i),
            "\t| Loss =",
            str(minibatch_loss),
            "\t| Accuracy =",
            str(minibatch_accuracy)
            )



# run on TEST images
test_accuracy = sess.run(accuracy, feed_dict={X: mnist.test.images, Y: mnist.test.labels, keep_prob: 1.0})
print("\nAccuracy on test set:", test_accuracy)



# try our own image
path = r"C:\Users\a.ibele\PycharmProjects\tensorflow\test_img.png"
img = np.invert(Image.open(path).convert('L')).ravel()

prediction = sess.run(tf.argmax(output_layer, 1), feed_dict={X: [img]})
print ("Prediction for test image:", np.squeeze(prediction))