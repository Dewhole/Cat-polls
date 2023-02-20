from .models import Users
from django.forms import ModelForm


class UsersForm(ModelForm):
    class Meta:
        model = Users
        fields = [
            "rating1",
            "unswer1",
            "rating2",
            "unswer2",
            "rating3",
            "unswer3",
            "rating4",
            "unswer4",
            "rating5",
            "unswer5",
            "rating6",
            "unswer6",
            "rating7",
            "unswer7",
            "rating8",
            "unswer8",
            "rating9",
            "unswer9",
            "rating10",
            "unswer10",
            "rating11",
            "unswer11",
            "rating12",
            "unswer12",
        ]
