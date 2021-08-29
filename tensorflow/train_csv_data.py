from __future__ import absolute_import, division, print_function
import os
import matplotlib.pyplot as plt
import tensorflow as tf
from tensorflow import contrib
# https://www.tensorflow.org/tutorials/eager/custom_training_walkthrough

tf.enable_eager_execution()
#print("TensorFlow version: {}".format(tf.__version__))
#print("Eager execution: {}".format(tf.executing_eagerly()))

train_dataset_fp = r"C:\Users\a.ibele\PycharmProjects\tensorflow\iris_training.csv"
column_names = ['sepal_length', 'sepal_width', 'petal_length', 'petal_width', 'species']     # column order in CSV file
feature_names = column_names[:-1]                                                            # = input (characteristics of flower)
label_name = column_names[-1]                                                                # = output (type of flower)
class_names = ['Iris setosa', 'Iris versicolor', 'Iris virginica']                           # names of outputs
batch_size = 32                                                                              # batch size = number of training samples worked through before internal parameters are updated

l_rate = 0.01                                                                                # learning rate. This is something to play around with a lot to get better results
optimizer = tf.train.GradientDescentOptimizer(learning_rate= l_rate)                         # specify optimizer

# create training dataset from csv
train_dataset = tf.contrib.data.make_csv_dataset(
    train_dataset_fp,
    batch_size,
    column_names=column_names,
    label_name=label_name,
    num_epochs=1) # epoch = number of complete passes through the dataset

features, labels = next(iter(train_dataset)) #next(iter( is just a demonstration of eager_execution since it displays something immediately

# plt.scatter(features['petal_length'].numpy(),
#             features['sepal_length'].numpy(),
#             c=labels.numpy(),
#             cmap='viridis')
#
# plt.xlabel("Petal length")
# plt.ylabel("Sepal length");

def pack_features_vector(features, labels): #Function to pack the features into a single array.
  features = tf.stack(list(features.values()), axis=1)
  return features, labels

train_dataset = train_dataset.map(pack_features_vector)  #Pack the features into a single array.

features, labels = next(iter(train_dataset)) #The features element of the Dataset are now arrays with shape (batch_size, num_features).

#build model
model = tf.keras.Sequential([
  tf.keras.layers.Dense(10, activation=tf.nn.relu, input_shape=(4,)),  # input shape required, 4 = number of identifying features in the dataset
  tf.keras.layers.Dense(10, activation=tf.nn.relu), #activation functions are requierd for hidden layers, relu is a common one
  tf.keras.layers.Dense(3)
])

predictions = model(features) # returns logits
tf.nn.softmax(predictions[:5]) # converts logits to probabilities for each class

# define loss function
def loss(model, x, y):
  y_ = model(x)
  return tf.losses.sparse_softmax_cross_entropy(labels=y, logits=y_)

l = loss(model, features, labels)

# define gradient function
def grad(model, inputs, targets):
  with tf.GradientTape() as tape:
    loss_value = loss(model, inputs, targets)
  return loss_value, tape.gradient(loss_value, model.trainable_variables)

global_step = tf.Variable(0)

loss_value, grads = grad(model, features, labels)

optimizer.apply_gradients(zip(grads, model.trainable_variables), global_step)



################# Train model
tfe = contrib.eager

# keep results for plotting
train_loss_results = []
train_accuracy_results = []

num_epochs = 201

for epoch in range(num_epochs):
    epoch_loss_avg = tfe.metrics.Mean()
    epoch_accuracy = tfe.metrics.Accuracy()

    # Training loop - using batches of 32
    for x, y in train_dataset:
        # Optimize the model
        loss_value, grads = grad(model, x, y)
        optimizer.apply_gradients(zip(grads, model.trainable_variables),
                                  global_step)

        # Track progress
        epoch_loss_avg(loss_value)  # add current batch loss
        # compare predicted label to actual label
        epoch_accuracy(tf.argmax(model(x), axis=1, output_type=tf.int32), y)

    # end epoch
    train_loss_results.append(epoch_loss_avg.result())
    train_accuracy_results.append(epoch_accuracy.result())

    if epoch % 50 == 0:
        print("Epoch {:03d}: Loss: {:.3f}, Accuracy: {:.3%}".format(epoch,
                                                                    epoch_loss_avg.result(),
                                                                    epoch_accuracy.result()))


# visualize performance of mode. Want to see loss go down and accuracy go up. These plots can help guide you to train the perfect model
fig, axes = plt.subplots(2, sharex=True, figsize=(12, 8))
fig.suptitle('Training Metrics')

axes[0].set_ylabel("Loss", fontsize=14)
axes[0].plot(train_loss_results)

axes[1].set_ylabel("Accuracy", fontsize=14)
axes[1].set_xlabel("Epoch", fontsize=14)
axes[1].plot(train_accuracy_results);




######################### test data
test_fp = 'C:\\Users\\a.ibele\\.keras\\datasets\\iris_test.csv'

test_dataset = tf.contrib.data.make_csv_dataset(
    test_fp,
    batch_size,
    column_names=column_names,
    label_name='species',
    num_epochs=1,
    shuffle=False)

test_dataset = test_dataset.map(pack_features_vector)

test_accuracy = tfe.metrics.Accuracy()

for (x, y) in test_dataset:
  logits = model(x)
  prediction = tf.argmax(logits, axis=1, output_type=tf.int32)
  test_accuracy(prediction, y)

print("Test set accuracy: {:.3%}".format(test_accuracy.result()))