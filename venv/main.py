
from tkinter import *
import sys
import imapclient
import time
import speech_recognition as sr
import pprint
import imaplib
import pyzmail
import email.message
import win32com.client as wincl



sr.Microphone.list_microphone_names()
root = Tk()
GUI = Toplevel()


def login(event):
    userEmailInput = emailFild.get()
    userPasswordInput = passwardFild.get()
    UserEmailProver = eProvider.get()
    domainName = StringVar

    if UserEmailProver == "Gmail":
        domainName = "imap.gmail.com"
    elif UserEmailProver == "Outlook":
        domainName = "imap-mail.outlook.com"
    elif UserEmailProver == "Yahoo":
        domainName = "imap.mail.yahoo.com"

    if emailFild.get() == userEmailInput and passwardFild.get() == userPasswordInput:
        imapObj = imapclient.IMAPClient(domainName, ssl=True)
        imapObj.login(userEmailInput, userPasswordInput)

        root.deiconify()
        GUI.destroy()


        def recognizeSpeech(recognizer, microphone):

            # check that recognizer and microphone arguments are working
            if not isinstance(recognizer, sr.Recognizer):
                raise TypeError("`recognizer` is not recognizer ")

            if not isinstance(microphone, sr.Microphone):
                raise TypeError("`microphone` not mic instance")

            # adjust the recognizer sensitivity to ambient noise and record audio
            # from the microphone
            with microphone as micIn:
                recognizer.adjust_for_ambient_noise(micIn)
                audio = recognizer.listen(micIn)

            # set up the feedback object
            feedback = {
                "success": True,
                "error": None,
                "log": None
            }

            try:
                feedback["log"] = recognizer.recognize_google(audio)
            except sr.RequestError:
                # API was unresponsive
                feedback["success"] = False
                feedback["error"] = "API unavailable"
            except sr.UnknownValueError:
                feedback["error"] = "Unable to recognize speech"

            return feedback

        if __name__ == "__main__":
            # get the email in box foulder list, give user chance to recover form failed inputs.
            emailFolders = 'inbox'
            User_Input_Attempts = 3
            User_Prompt = 2

            # create recognizer and mic instances
            recognizer = sr.Recognizer()
            microphone = sr.Microphone()

            word = emailFolders

            # format the instructions string
            userPrompt = (
                "Please ask whitch folder to be read:\n"
                "{words}\n"

            ).format(words=word)

            # show proprmt wait 1 sec befor
            print(userPrompt)
            time.sleep(1)

            for i in range(User_Input_Attempts):

                for j in range(User_Prompt):

                    input = recognizeSpeech(recognizer, microphone)
                    if input["log"]:
                        break
                    if not input["success"]:
                        break
                    print("I didn't catch that. What did you say?\n")

                # stop if eorror
                if input["error"]:
                    print("ERROR: {}".format(input["error"]))
                    break



                # compair user input to list of folders
                input_is_correct = input["log"] == word
                user_has_more_attempts = i < User_Prompt - 1

                # Open the folder that user has requested
                # list emails within that foulder

                if input_is_correct:
                    print("Opening inbox".format(word))
                    imapObj.select_folder(word, readonly=True)
                    uid = imapObj.search(
                        ['ALL'])  # This will retuen all emails within inbox(for unread emails change to UNSEEN
                    print(uid)
                    break
                elif user_has_more_attempts:
                    print("Incorrect. Try again.\n")
                else:
                    print("Sorry, i am looking for'{}'.".format(word))
                    break

        """""
        This is pulling the raw message file 
        from the IMAP server in RFC 822 format
        will be made readable by pyzmail lib.

        """""
        rawMessage = imapObj.fetch(uid, ['BODY[]'])
        # pprint.pprint(rawMessage)

        message = pyzmail.PyzMessage.factory(rawMessage[10][b'BODY[]'])
        messageAdd = message.get_address('from')
        messageSub = message.get_subject()
        messageBod = message.text_part.get_payload().decode(message.html_part.charset)

        print(messageAdd)
        print(messageSub)
        print(messageBod)

        speak = wincl.Dispatch("SAPI.SpVoice")
        emailFile = open("testEmail.txt", "a+")
        mail = [messageAdd, messageSub, messageBod]
        emailFile.write(str(mail))

        emailFile = open("testEmail.txt", "r")
        if emailFile.mode == "r":
            contents = emailFile.read()
            print(contents)
            speak.Speak(contents)

def quit_application():
    GUI.destroy()
    root.destroy()
    sys.exit()

emailProvider = [
    "Gmail",
    "Outlook",
    "Yahoo",
]
eProvider = StringVar()
eProvider.set(emailProvider[0])


GUI.geometry("500x250")
GUI.title("Talking Email")
GUI.configure(background="gray")
usernameLbl = Label(GUI, text="Email Adress", font=("Helvetica", 11))
emailFild = Entry(GUI)
passwardLbl = Label(GUI, text="Password", font=("Helvetica", 11))
passwardFild = Entry(GUI, show="*")
emailProviderDrop = OptionMenu(GUI, eProvider, *emailProvider)
loginBtn = Button(GUI, text="Login", command=lambda:login)
quitBtn = Button(GUI, text="Exit", command=lambda:quit_application())


passwardFild.bind("<Return>", login)

usernameLbl.pack()
emailFild.pack()
passwardLbl.pack()
passwardFild.pack()
emailProviderDrop.pack()
loginBtn.pack()
quitBtn.pack()

root.title("Talking Email")
root.configure(background="gray")
root.geometry("500x250")
displayMessage = Label(root, text="You have logged in try saying inbox for new messages", font=("Helvetica", 11))
displayMessage.pack


root.withdraw()
root.mainloop()









