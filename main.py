# main.py

import click
from automation import lab1, lab2, lab3, lab4, lab5, lab6, lab7

@click.command()
@click.option('--lab', type=click.Choice(['1','2','3','4','5','6','7']), required=True)
def main(lab):
    module = {'1': lab1, '2': lab2, '3': lab3, '4': lab4,
              '5': lab5, '6': lab6, '7': lab7}[lab]
    module.run()

if __name__ == '__main__':
    main()
