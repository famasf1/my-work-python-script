from email.policy import default
import firebase_admin as fb
from firebase_admin import firestore
import pandas as pd


cred = fb.credentials.Certificate(rf"C:\ITEC_Support\my-work-python-script\Database_in_pandas\serviceAccount.json")
default_app = fb.initialize_app(cred)
db = firestore.client()

my_doc_Ref = db.collection("Data").document("ID:33_Insure")
my_doc_Ref.set()