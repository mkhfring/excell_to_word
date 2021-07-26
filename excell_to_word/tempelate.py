# mode defines as follow: 1 simple paragraph
# 2 simple sentece1
# 3 bold
# 4 bold_paragraph
# 5 under_line
# 6 paragraph with parameters
from enum import Enum


class SentenceElement(Enum):
    PARAGRAPH = 1
    BOLD_PARAGRAPH = 2
    SENTENCE = 3
    UNDER_LINE = 4
    PRAGRAPH_PARAM = 5
    BOLD = 6


name = "Mohamad"
template = [
(
     """
The above hours are an average per week, for 7 weeks (1 week of which is during the final exam period).  Hours worked in any one given week may be more or less than the above average. Please note that this is a salaried position and it is the responsibility of the TA and course instructor to manage working hours. If additional hours are required to complete a task, they must be pre-approved by the CMPS Department. Any hours submitted through Workday that was not approved prior will be denied.
""", SentenceElement.PARAGRAPH
),
(
    """Note that this assignment extends throughout, and possibly beyond, the final exam period """, SentenceElement.BOLD_PARAGRAPH
),

( 
    """and you may be asked to mark final exams at a date that is""", SentenceElement.SENTENCE
),

(
    """after the last day final exams are being written.""", SentenceElement.BOLD
),

(
    """By accepting this position you acknowledge that you are willing to stay for the full contracted work term, which includes the final exam period""", SentenceElement.SENTENCE
),

(
    """If you accept this offer, please reply to Chad Davis via e-mail at cdavis.cmpsta@ubc.ca by """, SentenceElement.PARAGRAPH
),

(
    """August 10, 2021. """, SentenceElement.BOLD
),

(
    """If you can let us know sooner it would be greatly appreciated.""", SentenceElement.UNDER_LINE
),

(
    """This assignment is not yet finalized and may be subject to change depending on course enrolment figures and TA availability. If necessary you will be contacted regarding any changes, so please keep checking your e-mail
""", SentenceElement.PARAGRAPH
),

(
     """Kind regards,
""", SentenceElement.PARAGRAPH
),

(
    """Dr. Chad Davis
""", SentenceElement.BOLD_PARAGRAPH
),

(
    """
Lecturer
IKBSAS Department of Computer Science, Math, Physics, & Statistics
The University of British Columbia | Okanagan Campus

""", SentenceElement.SENTENCE
),
]

if __name__ == "__main__":
    print(template)