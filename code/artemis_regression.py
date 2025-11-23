import pandas as pd
from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.preprocessing import StandardScaler, PolynomialFeatures
from sklearn.pipeline import Pipeline
from sklearn.linear_model import Ridge
from sklearn.metrics import r2_score

df = pd.read_excel("../data/artemis/artemis_data_for_regression.xlsx")

X = df.select_dtypes("number").drop(columns=["Rating"])
y = df["Rating"]

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.33, random_state=42)

model = Pipeline([
    ("poly", PolynomialFeatures(degree=2, include_bias=False)),
    ("scaler", StandardScaler()),
    ("ridge", Ridge(alpha=10))
])

cv_r2 = cross_val_score(model, X, y, cv=5, scoring="r2").mean()

model.fit(X_train, y_train)
test_r2 = r2_score(y_test, model.predict(X_test))

print("CV R²:", cv_r2)
print("Test R²:", test_r2)
