from pydantic import BaseModel, Field


class Transaction(BaseModel):
    card: str = Field(alias = "Credit Card")
    merchant: str = Field(alias = "Merchant")
    amount: str = Field(alias = "Paid Amount")
    date: str = Field(alias = "Date")

    def get_object(self):
        return self.model_dump(exclude_none=True)