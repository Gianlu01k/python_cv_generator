from docx import Document
from docx.shared import Inches
import pyttsx3

def speak(text):
    pyttsx3.speak(text)

document = Document()

#profile picture
document.add_picture(
    'profile_picture.jpg',
    width=Inches(2.0)
)


#name, phone number and email
speak('Ciao, qual è il tuo nome?')
firstname = input('Ciao, qual è il tuo nome? ')
speak(firstname +', qual è il tuo cognome? ')
lastname = input(firstname +', qual è il tuo cognome? ')

name=firstname + ' ' + lastname

speak('Qual''è il tuo numero di telefono?')
phone_number = input('Qual''è il tuo numero di telefono? ')

speak('Qual''è la tua mail?')
email = input('Qual''è la tua mail? ')

document.add_paragraph(
    name + ' | ' + phone_number + ' | ' + email
)

#about me
document.add_heading('Chi sono')
speak('Parlami di te')
about_me = input('Parlami di te ')
document.add_paragraph(about_me)

#work experience
document.add_heading('Esperienze lavorative')
p = document.add_paragraph()

speak('In quale azienda hai lavorato?')
company = input('Inserisci azienda ')
speak('Da quando hai lavorato in '+company+'?')
from_date = input('Dalla data ')
speak('Fino a quando?')
to_date = input('Alla data ')

p.add_run(company + ' ').bold = True
p.add_run(from_date + '-' + to_date + '\n').italic = True

speak('Descrivi la tua esperienza')
experience_details = input(
    'Descrivi la tua esperienza in ' + company)
p.add_run(experience_details)

#skills
document.add_heading('Competenze')
speak('Quali competenze hai?')
skill = input('Inserisci competenza ')
p = document.add_paragraph(skill)
p.style = 'List Bullet'

while True:
    speak('Vuoi inserire un''altra competenza?')
    q = input('Vuoi inserire un''altra competenza? si/no ')
    if q.lower() == 'si' or q.lower() == 's':
        speak('Quali competenze hai?')
        skill = input('Inserisci competenza ')
        p = document.add_paragraph(skill)
        p.style= 'List Bullet'
    else:
        break

#more experiences
while True:
    speak('Hai altre esperienze lavorative?')
    has_more_exp = input('Hai altre esperienze lavorative? si/no ')
    if has_more_exp.lower == 'si' or has_more_exp.lower() == 's':
        speak('In quale azienda hai lavorato?')
        company = input('Inserisci azienda ')
        speak('Da quando hai lavorato in '+company+'?')
        from_date = input('Dalla data ')
        speak('Fino a quando?')
        to_date = input('Alla data ')

        p.add_run(company + ' ').bold = True
        p.add_run(from_date + '-' + to_date + '\n').italic = True

        speak('Descrivi la tua esperienza')
        experience_details = input(
          'Descrivi la tua esperienza in ' + company)
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
