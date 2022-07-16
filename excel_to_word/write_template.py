import os
import glob
import threading

import click

from excel_to_word.letter import OfficialLetter, OfferLetter
from excel_to_word.exam_letter import ExamLetter, MarkerLetter

HERE = os.path.dirname(os.path.realpath(__file__))
paragraphs = []


def delete_previous_files(path):
    dirs = os.listdir(path)
    for element in dirs:
        new_path = os.path.join(path, element)
        if os.path.isfile(new_path):
            if new_path.endswith(".xlsx"):
                continue

            os.remove(new_path)
        if os.path.isdir(new_path):
            delete_previous_files(new_path)




@click.command()
@click.option('--path', default=os.path.join(HERE, "data/main_data.xlsx"),
              help='number of greetings')
def main(path):
    print(f"The inserted path is {path}")

    delete_previous_files(os.path.join(HERE, "data"))
    official_letter = OfficialLetter(
        path,
        os.path.join(HERE, "templates/letters.docx"),
        os.path.join(HERE, "templates/letters_temp.docx")
        #os.path.join(HERE, "templates/output.docx")

    )
    #t1 = threading.Thread(target=official_letter.create_output)
    #t1.start()
    official_letter.create_output()
    offer_letter = OfferLetter(
        path,
        os.path.join(HERE, "templates/offer.docx"),
    )
    #t = threading.Thread(target=offer_letter.create_output)
    #t.start()
    offer_letter.create_output()

@click.command()
@click.option('--path', default=os.path.join(HERE, "data/main_data.xlsx"),
              help='number of greetings')

def letter(path):
    delete_previous_files(os.path.join(HERE, "data"))
    marker_letter = MarkerLetter(
        path,
        os.path.join(HERE, "templates/students.docx"),
    )
    marker_letter.create_output()
    print("Hello")


if __name__ == "__main__":
     # main()
     letter()