from docx import Document
from docx.shared import Inches
import pyttsx3

def speak(text):
    pyttsx3.speak(text)

document = Document()

#profile picture
#document.add_picture(
    #'path',
    # width=Inces(dim.0) 
#)


#name, phone number and email
name = input('What is your name? ')
speak('Hello '+ name + 'how are you today?')

speak('What is your phone number? ')
phone_number = input()
email = input('What is your email? ')

document.add_paragraph(
    name + ' | ' + phone_number + ' | ' + email
)

#about me
document.add_heading('About me')
about_me = input('Tell me about yourself ')
document.add_paragraph(about_me)

#work experience
document.add_heading('Work Experience')
p = document.add_paragraph()

company = input('Enter company ')
from_date = input('From Date')
to_date = input('To Date')

p.add_run(company + ' ').bold = True
p.add_run(from_date + '-' + to_date + '\n').italic = True

experience_details = input(
    'Describe your experience at ' + company)
p.add_run(experience_details)

#skills
document.add_heading('Skills')
skill = input('Insert skill: ')
p = document.add_paragraph(skill)
p.style = 'List Bullet'

while True:
    q = input('Do you want to add another skill ? Yes/No')
    if q.lower() == 'yes' or q.lower() == 'y':
        skill = input('Insert skill: ')
        p = document.add_paragraph(skill)
        p.style= 'List Bullet'
    else:
        break

#more experiences
while True:
    has_more_exp = input('Do you have more work experiences? ')
    if has_more_exp.lower == 'yes':
        company = input('Enter company ')
        from_date = input('From Date')
        to_date = input('To Date')

        p.add_run(company + ' ').bold = True
        p.add_run(from_date + '-' + to_date + '\n').italic = True

        experience_details = input(
        'Describe your experience at ' + company)
        p.add_run(experience_details)
    else:
        break

#footer
#section = document.section[0]
#footer = section.footer
#p = footer.paragraph[0]
#p.text = "Cv generated with python"
#
#doc saving
document.save('cv.docx')