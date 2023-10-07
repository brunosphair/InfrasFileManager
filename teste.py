def soma(x):
    x = x + 1
    return x


def acresentar(lista):
    lista.append('teste2')


def deletar(dict):
    del dict['deletar']


if __name__ == '__main__':
    dict = {}
    dict['deletar'] = 'merda'
    deletar(dict)
    print(dict)
