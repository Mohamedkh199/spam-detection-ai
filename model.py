import pandas as pd
import pickle
from sklearn.model_selection import train_test_split
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB

# تحميل الداتا
data = pd.read_csv("spam.csv", encoding="latin-1")

# تنظيف الأعمدة
data = data[['v1','v2']]
data.columns = ['label','message']

data['label'] = data['label'].map({'spam':1, 'ham':0})

# تقسيم البيانات
X_train, X_test, y_train, y_test = train_test_split(
    data['message'], data['label'], test_size=0.2, random_state=42)

# تحويل النصوص لأرقام
vectorizer = TfidfVectorizer()
X_train_vec = vectorizer.fit_transform(X_train)

# تدريب الموديل
model = MultinomialNB()
model.fit(X_train_vec, y_train)

# حفظ الموديل والـ vectorizer
pickle.dump(model, open("model.pkl", "wb"))
pickle.dump(vectorizer, open("vectorizer.pkl", "wb"))

print("Model trained and saved successfully!")
