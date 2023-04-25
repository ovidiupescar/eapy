import win32com.client
import sys
import openai

API_KEY = "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"

def chat_gpt(prompt):
    """
    Sends a promt to Chat GPT API and returns the answer
    """

    print(prompt)

    openai.api_key = API_KEY
    
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=0.5,
        max_tokens=1000
        )

    return response["choices"][0]["text"]


def prepare(response):
    """
    Splits the ChatGPT response string into different testcases (based on endline)
    """
    test_cases = response.split('\n')
    return test_cases

def clear_tests(req):
    """
    Clear tests that are already connected to the selected requirement
    """

    tests = req.Tests
    
    for test in range(0, tests.Count):
        current_test = tests.GetAt(test)
        print(current_test.name)
        tests.DeleteAt(test, False)


def add_test(req, test, test_type):
    """
    Adds a test to the requirement
    """
    tests = req.Tests
    newTest = tests.AddNew( test, test_type)
    newTest.Update()
    tests.Refresh()



try:
    eaApp = win32com.client.Dispatch("EA.App")
except:
    sys.stderr.write( "Unable to connect to EA\n" )
    exit()


mEaRep = eaApp.Repository

if mEaRep.ConnectionString == '':
    sys.stderr.write( "EA has no Model loaded\n" )
    exit()


print("connecting...", mEaRep.ConnectionString)


selected_req = mEaRep.GetContextObject()

clear_tests(selected_req)

prompt = "Define test cases for the requirement: "

req_descr = f"{selected_req.Name}: {selected_req.Notes}"

response = chat_gpt(prompt+req_descr)

test_cases = prepare(response)

for test_case in test_cases:
    add_test(selected_req, test_case, "System")
print("FINISHED!")
    