from django import forms
from .models import Order

class OrderForm(forms.ModelForm):
    class Meta:
        model = Order
        fields = [
            "branch",
            "client_name",
            "item_name",
            "service_type",
            "defect_description",
            "desired_date",
            "urgent",
        ]