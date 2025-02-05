#IMPORT
from keras.models import load_model
from PIL import Image
import numpy as np
import time
import matplotlib.pyplot as plt
import cv2
import sys
from visualizer import visualize
from sklearn.metrics import confusion_matrix, plot_confusion_matrix
import itertools

"""
# Classe permettant de réaliser une prédiction sur une nouvelle donnée
"""


def main():
    """
    # On definit les chemins d'acces au différentes hyper parametre
    """

    modelPath = 'C:\\model.h5'
    imagePath = '.\\predict\\test1.jpg'
    maskPath = '.\\predict\\mask1.png'

    #predict(modelPath, imagePath)
    predictNconfusion(modelPath, imagePath, maskPath)


def predictNconfusion(modelPath, imagePath, maskPath):
    image = Image.open(imagePath).convert('RGB')
    img = image.resize(size=(256, 256))
    img = np.asarray(img, dtype=np.float32) / 255.
    print("START LOAD")
    model = load_model(modelPath, compile=False)
    print("END LOAD")

    dimension = img.shape
    img = img.reshape(1, dimension[0], dimension[1], dimension[2])

    prediction = model.predict(img)
    res = np.asarray(prediction[0]*100)
    res[res >= 0.95] = 1
    res[res < 0.95] = 0
    np.set_printoptions(threshold=sys.maxsize)
    mask = Image.open(maskPath)
    mask = mask.resize(size=(256, 256))

    maskNp = np.asarray(mask)
    #print(maskNp)
    #visualize(image, mask)

    res = res[:, :, 0]
    print(maskNp.shape)
    print(res.shape)
    print(maskNp.dtype)
    print(res.dtype)
    res = res.astype(np.uint8)

    test = confusion_matrix(maskNp.flatten(), res.flatten())
    test = test.astype('float') / test.sum(axis=1)[:, np.newaxis]
    #plot_confusion_matrix(clf, maskNp.flatten(), maskNp.flatten())

    plt.figure()
    cmap = plt.cm.Blues
    classes = ['background', 'carrie']
    title = 'Confusion matrix'
    plt.imshow(test, interpolation='nearest', cmap=cmap)
    plt.title(title)
    plt.colorbar()
    tick_marks = np.arange(2)
    plt.xticks(tick_marks, classes, rotation=45)
    plt.yticks(tick_marks, classes)

    fmt = '.2f'
    thresh = test.max() / 2.
    for i, j in itertools.product(range(test.shape[0]), range(test.shape[1])):
        plt.text(j, i, format(test[i, j], fmt), horizontalalignment="center",
                 color="white" if test[i, j] > thresh else "black")

    plt.ylabel('True label')
    plt.xlabel('Predicted label')
    plt.tight_layout()
    plt.show()
    plt.imshow(res)
    plt.show()


def predict(modelPath,imagePath):
    """
    # Fonction qui permet de convertir une image en array, de charger le modele et de lui injecter notre image pour une prediction
    :param modelPath: chemin du modèle au format hdf5
    :param imagePath: chemin de l'image pour realiser une prediction
    :param imageSize: défini la taille de l'image. IMPORTANT : doit être de la même taille que celle des images
    du dataset d'entrainements
    :param label: nom de nos 5 classes de sortie
    """

    start = time.time()

    # Chargement du modele
    print("Chargement du modèle :\n")
    model = load_model(modelPath)
    print("\nModel chargé.")

    #Chargement de notre image et traitement
    data = []
    img = Image.open(imagePath).convert('RGB')
    img = img.resize(size=(256, 256))
    #img.load()
    #img = img.resize(size=imageSize)
    img = np.asarray(img, dtype=np.float32) / 255.
    #img = np.asarray(img)
    #data.append(img)
    #data = np.asarray(data)
    plt.imshow(img)
    plt.show()

    #On reshape pour correspondre aux dimensions de notre modele
    # Arg1 : correspond au nombre d'image que on injecte
    # Arg2 : correspond a la largeur de l'image
    # Arg3 : correspond a la hauteur de l'image
    # Arg4 : correspond au nombre de canaux de l'image (1 grayscale, 3 couleurs)
    #dimension = data[0].shape
    dimension = img.shape
    print(dimension)
    #Reshape pour passer de 3 à 4 dimension pour notre réseau
    #data = data.astype(np.float32).reshape(data.shape[0], dimension[0], dimension[1], dimension[2])
    img = img.reshape(1, dimension[0], dimension[1], dimension[2])
    np.set_printoptions(threshold=sys.maxsize)
    #On realise une prediction
    prediction = model.predict(img)
    res = np.asarray(prediction[0]*100)
    print("PREDICTION\n")
    print(res)
    print(res.shape)
    # MULTICLASS
    #res = np.argmax(res, axis = 2)
    res[res >= 0.95] = 255
    res[res <= 0.1] = 0


    #print(res)
    #print(res.shape)


    plt.imshow(res)
    plt.show()

    '''
    pr_mask = model.predict(np.expand_dims(img, axis=0)).squeeze()
    # pr_mask.shape == (H, W, C)

    pr_mask = np.argmax(pr_mask, axis=2)
    # pr_mask.shape == (H, W)

    # to count the occurrences, say car is equal to 3 in the pr_mask
    num_car_pixels = numpy.count_nonzero(pr_mask == 3)
    percent_car_pixels = (num_car_pixels / (H * W)) * 100
    '''







    '''
    test11 = np.asarray(prediction[0], dtype=np.float32)
    test22 = np.asarray(prediction[0], dtype=np.uint8)
    test1 = np.asarray(prediction[0]*100, dtype=np.float32)


    test2 = np.asarray(prediction[0]*100, dtype=np.uint8)
    test3 = np.asarray(prediction[0]*255, dtype=np.uint8)
    test4 = np.asarray(prediction[0]*255, dtype=np.float32)



    cv2.imwrite("predict/predicted.jpg", cv2.cvtColor(prediction[0], cv2.COLOR_RGB2BGR))
    cv2.imwrite("predict/test2.jpg", cv2.cvtColor(test2, cv2.COLOR_RGB2BGR))
    cv2.imwrite("predict/test3.jpg", cv2.cvtColor(test3, cv2.COLOR_RGB2BGR))
    cv2.imwrite("predict/test4.jpg", cv2.cvtColor(test4, cv2.COLOR_RGB2BGR))

    plt.imshow(cv2.cvtColor(prediction[0], cv2.COLOR_RGB2BGR))
    plt.show()
    plt.imshow(cv2.cvtColor(test1, cv2.COLOR_RGB2BGR))

    plt.show()
    plt.imshow(test2)
    plt.show()
    plt.imshow(test3)
    plt.show()
    plt.imshow(test4)
    plt.show()

    plt.imshow(test11)
    plt.show()
    plt.imshow(test22)
    plt.show()

    class_index = np.argmax(prediction, axis=2)
    colors = {0: [255, 255, 255]}
    colored_image = np.array([colors[x] for x in np.nditer(class_index)],
                             dtype=np.uint8)
    output_image = np.reshape(colored_image, (256, 256, 3))
    plt.imshow(output_image)
    plt.show()
    '''
    #On recupere le mot correspondant à l'indice precedent
    #word = label[maxPredict]
    #pred = prediction[0][maxPredict] * 100.
    end = time.time()


    #On affiche les prédictions
    print()
    print('----------')
    print(" Prediction :")
    print('temps prediction : ' + "{0:.2f}secs".format(end-start))

    print('----------')


if __name__ == "__main__":
    """
    # MAIN
    """
    main()
