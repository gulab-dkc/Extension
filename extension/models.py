from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone
# Create your models here.


class Employee(models.Model):
    """
    Employee model representing basic information about an employee.

    Fields:
        name (CharField): The full name of the employee, with a maximum length of 100 characters.
        email (EmailField): The email address of the employee. Must be unique for proper identification.
        is_active (BooleanField): Indicates whether the employee is currently active. Defaults to True.
        created_at (DateTimeField): Timestamp when the employee record was created. Defaults to the current time.
        updated_at (DateTimeField): Timestamp when the employee record was last updated. Defaults to the current time.

    Methods:
        __str__(): Returns the string representation of the employee, which is the employee's name.
    """
    name = models.CharField(max_length=100)
    email = models.EmailField()
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(default=timezone.now)
    updated_at = models.DateTimeField(default=timezone.now)

    def __str__(self):
        return self.name



class Extension(models.Model):
    """
    Model to store user extension details.

    Fields:
    - user: One-to-one relationship with Django's built-in User model.
    - extension: String field to store the extension number (max length 5).
    - status: Boolean field to indicate whether the extension is active.
    - created_at: Date and time when the extension was created.
    - updated_at: Date and time when the extension was last updated.
    """
    emp = models.OneToOneField(Employee, on_delete=models.CASCADE)
    extension = models.CharField(max_length=5)
    status = models.BooleanField(default=False)
    created_at = models.DateTimeField(default=timezone.now)
    updated_at = models.DateTimeField(default=timezone.now)

    def __str__(self):
        return f"{self.emp.name} - {self.extension}"





