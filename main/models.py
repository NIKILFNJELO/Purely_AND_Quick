from django.db import models
from django.contrib.auth.models import User
from django.db.models.signals import post_save
from django.dispatch import receiver


class Profile(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='profile')
    phone = models.CharField(max_length=20, blank=True, default='')
    loyalty_status = models.CharField(max_length=50, default='Звичайний')
    avatar = models.ImageField(upload_to='avatars/', blank=True, null=True)
    paid_orders_count = models.PositiveIntegerField(default=0)

    def __str__(self):
        return self.user.username


class Order(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    branch = models.CharField(max_length=50)
    item_name = models.CharField(max_length=100)
    service_type = models.CharField(max_length=200)
    desired_date = models.DateField()
    urgent = models.BooleanField(default=False)
    is_paid = models.BooleanField(default=False)
    total_price = models.DecimalField(max_digits=8, decimal_places=2, default=0)
    defect_description = models.TextField(blank=True)
    order_date = models.DateField(auto_now_add=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"Замовлення {self.item_name} від {self.user.username}"


class Review(models.Model):
    author = models.ForeignKey(User, on_delete=models.CASCADE, related_name="reviews")
    title = models.CharField(max_length=200)
    text = models.TextField()
    rating = models.PositiveSmallIntegerField(default=5)
    is_company = models.BooleanField(default=False)
    company = models.CharField(max_length=100, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-created_at"]

    def __str__(self):
        return f"{self.author.username} — {self.title} ({self.rating}★)"


@receiver(post_save, sender=User)
def create_profile(sender, instance, created, **kwargs):
    if created:
        Profile.objects.get_or_create(user=instance, defaults={'phone': ''})