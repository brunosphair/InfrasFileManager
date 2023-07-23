from dotenv import dotenv_values

try:
    config = dotenv_values('config.env')
    print(config)
    print(config['REGULAR_EXPRESSION'])
except KeyError:
    pass