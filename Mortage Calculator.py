def calculate_mortgage(principal, interest_rate, years):
    
    # Calculate the monthly mortgage payment given the principal amount, annual interest rate, and loan term in years.
    
    monthly_interest_rate = interest_rate / 100 / 12
    total_payments = years * 12

    # Calculate the monthly mortgage payment
    mortgage_payment = (principal * monthly_interest_rate) / (1 - (1 + monthly_interest_rate) ** -total_payments)

    return mortgage_payment


def mortgage_schedule(principal, interest_rate, years):
    
    # Generate a mortgage payment schedule showing the monthly payment, interest paid, principal paid, and remaining balance.
    
    monthly_payment = calculate_mortgage(principal, interest_rate, years)
    remaining_balance = principal

    print("Mortgage Payment Schedule:")
    print("--------------------------")
    print("Month\tPayment\tInterest\tPrincipal\tBalance")

    month = 1
    while month <= years * 12:
        interest = remaining_balance * interest_rate / 100 / 12
        principal_payment = monthly_payment - interest
        remaining_balance -= principal_payment

        #Print mortgage schedule to two decimals places
        print(str(month) + "\t" + format(monthly_payment, ".2f") + "\t" + format(interest, ".2f") + "\t\t" + format(principal_payment, ".2f") + "\t\t" + format(remaining_balance, ".2f"))

        month += 1


# Example usage
principal = 100000
interest_rate = 10
years = 25

monthly_payment = calculate_mortgage(principal, interest_rate, years)
total_payment = monthly_payment * years * 12

print("Principal amount: $" + str(principal))
print("Interest rate: " + str(interest_rate) + "%")
print("Loan term: " + str(years) + " years")
print("Monthly mortgage payment: $" + format(monthly_payment, ".2f"))
print("Total payment: $" + format(total_payment, ".2f"))

mortgage_schedule(principal, interest_rate, years)


