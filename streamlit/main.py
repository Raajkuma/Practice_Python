from fastapi import FastAPI, Query

app = FastAPI(title="My First API")


@app.get("/")
def read_root():
    return {"message": "Welcome to my first FastAPI app!"}


# Path parameter: the value comes from the URL itself
@app.get("/greet/{name}")
def greet(name: str):
    return {"message": f"Hello, {name}! Welcome to FastAPI."}


# Query parameter: the value comes after the ? in the URL
@app.get("/loan-interest")
def loan_interest(
    amount: float = Query(..., gt=0, description="Loan amount"),
    rate: float = Query(..., ge=0, description="Annual interest rate in %"),
    years: float = Query(..., gt=0, description="Duration in years"),
):
    # Simple interest: I = P * R * T / 100
    interest = amount * rate * years / 100
    total_paid = amount + interest

    return {
        "loan_amount": amount,
        "interest_rate_percent": rate,
        "duration_years": years,
        "total_interest": round(interest, 2),
        "total_amount_paid": round(total_paid, 2),
    }