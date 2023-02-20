from django.db import models
from datetime import date

from django.urls import reverse


class Questions(models.Model):
    title = models.CharField("Тема", max_length=100)
    question1 = models.TextField("Вопрос 1", blank=True, null=True)
    question2 = models.TextField("Вопрос 2", blank=True, null=True)
    question3 = models.TextField("Вопрос 3", blank=True, null=True)
    question4 = models.TextField("Вопрос 4", blank=True, null=True)
    question5 = models.TextField("Вопрос 5", blank=True, null=True)
    question6 = models.TextField("Вопрос 6", blank=True, null=True)
    question7 = models.TextField("Вопрос 7", blank=True, null=True)
    question8 = models.TextField("Вопрос 8", blank=True, null=True)
    question9 = models.TextField("Вопрос 9", blank=True, null=True)
    question10 = models.TextField("Вопрос 10", blank=True, null=True)
    question11 = models.TextField("Вопрос 11", blank=True, null=True)
    question12 = models.TextField("Вопрос 12", blank=True, null=True)

    def __str__(self):
        return self.title

    class Meta:
        verbose_name = "Анкеты Transactional Survey"
        verbose_name_plural = "Анкеты Transactional Survey"


class Users(models.Model):

    name = models.CharField("Имя", max_length=100)
    company = models.CharField("Компания", max_length=100, blank=True, null=True)
    anketa = models.CharField("Анкета", max_length=100, blank=True, null=True)
    listPhone = models.CharField(
        "Список обзвона", max_length=100, blank=True, null=True
    )
    datePhone = models.CharField("Дата звонка", max_length=100, blank=True, null=True)
    timePhone = models.CharField("Время звонка", max_length=100, blank=True, null=True)
    timesPhone = models.CharField(
        "Длительность звонка (в минутах)", max_length=100, blank=True, null=True
    )
    lucky = models.CharField("Звонок успешен", max_length=100, blank=True, null=True)
    partyID = models.CharField("Код каталога", max_length=100, null=True, blank=True)
    description = models.TextField("Описание (Вы купили)", blank=True, null=True)
    facture = models.CharField("Фактура", max_length=200, blank=True, null=True)
    factureDate = models.CharField(
        "Дата фактуры", max_length=100, blank=True, null=True
    )
    phone = models.CharField("Телефон", max_length=100, blank=True, null=True)
    phoneCompany = models.CharField(
        "Телефон Организации", max_length=100, blank=True, null=True
    )
    phoneMobile = models.CharField(
        "Мобильный Телефон", max_length=100, blank=True, null=True
    )
    mail = models.EmailField(blank=True, null=True)
    loadDate = models.DateField("Дата загрузки", default=date.today)
    url = models.SlugField(max_length=130, unique=True, blank=True)
    title = models.ForeignKey(
        Questions,
        verbose_name="Тема опроса",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
    )
    rating1 = models.PositiveIntegerField("Ответ 1", blank=True, null=True)
    unswer1 = models.TextField("Примечание 1", blank=True, null=True)
    rating2 = models.PositiveIntegerField("Ответ 2", blank=True, null=True)
    unswer2 = models.TextField("Примечание 2", blank=True, null=True)
    rating3 = models.PositiveIntegerField("Ответ 3", blank=True, null=True)
    unswer3 = models.TextField("Примечание 3", blank=True, null=True)
    rating4 = models.PositiveIntegerField("Ответ 4", blank=True, null=True)
    unswer4 = models.TextField("Примечание 4", blank=True, null=True)
    rating5 = models.PositiveIntegerField("Ответ 5", blank=True, null=True)
    unswer5 = models.TextField("Примечание 5", blank=True, null=True)
    rating6 = models.PositiveIntegerField("Ответ 6", blank=True, null=True)
    unswer6 = models.TextField("Примечание 6", blank=True, null=True)
    rating7 = models.PositiveIntegerField("Ответ 7", blank=True, null=True)
    unswer7 = models.TextField("Примечание 7", blank=True, null=True)
    rating8 = models.PositiveIntegerField("Ответ 8", blank=True, null=True)
    unswer8 = models.TextField("Примечание 8", blank=True, null=True)
    rating9 = models.PositiveIntegerField("Ответ 9", blank=True, null=True)
    unswer9 = models.TextField("Примечание 9", blank=True, null=True)
    rating10 = models.PositiveIntegerField("Ответ 10", blank=True, null=True)
    unswer10 = models.TextField("Примечание 10", blank=True, null=True)
    rating11 = models.PositiveIntegerField("Ответ 11", blank=True, null=True)
    unswer11 = models.TextField("Примечание 11", blank=True, null=True)
    rating12 = models.PositiveIntegerField("Ответ 12", blank=True, null=True)
    unswer12 = models.TextField("Примечание 12", blank=True, null=True)

    def __str__(self):
        return self.name

    def get_absolute_url(self):
        return reverse("individualurl", kwargs={"slug": self.url})

    class Meta:
        verbose_name = "Результаты опросов"
        verbose_name_plural = "Результаты опросов"
