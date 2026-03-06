import json
import re
import logging

from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.models import User
from django.contrib.auth import authenticate, login, logout
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.views.decorators.http import require_POST, require_http_methods
from django.views.generic import TemplateView
from django.contrib.auth.mixins import LoginRequiredMixin
from django.http import JsonResponse

from .models import Profile, Order, Review

logger = logging.getLogger(__name__)


class HomePageView(TemplateView):
    template_name = 'main/about_us.html'

class ProposalsView(TemplateView):
    template_name = 'main/proposals.html'


@login_required
def show_more(request):
    profile = request.user.profile
    has_discount = profile.paid_orders_count >= 3
    orders_until_discount = max(0, 3 - profile.paid_orders_count)
    paid = profile.paid_orders_count
    progress_width = min(100, int((paid / 3) * 100))

    all_orders = Order.objects.filter(user=request.user)
    total_orders = all_orders.count()
    total_spent = sum(o.total_price for o in all_orders if o.is_paid)
    recent_orders = all_orders.order_by('-id')[:5]

    return render(request, "main/show_more.html", {
        "has_discount": has_discount,
        "orders_until_discount": orders_until_discount,
        "total_orders": total_orders,
        "total_spent": total_spent,
        "recent_orders": recent_orders,
        "progress_width": progress_width,
        "paid": paid,
    })


def catalog(request):
    has_discount = False
    if request.user.is_authenticated:
        profile, _ = Profile.objects.get_or_create(user=request.user, defaults={"phone": ""})
        has_discount = profile.paid_orders_count >= 3
    return render(request, "main/catalog.html", {
        "has_discount": has_discount,
        "user_authenticated": request.user.is_authenticated,  
    })

@login_required
def profile(request):
    if request.method == "POST" and request.FILES.get("avatar"):
        request.user.profile.avatar = request.FILES["avatar"]
        request.user.profile.save()
        return redirect("profile")

    orders = Order.objects.filter(user=request.user).order_by('id')
    for i, order in enumerate(orders, start=1):
        order.user_number = i

    orders_until_discount = max(0, 3 - request.user.profile.paid_orders_count)

    return render(request, "main/profile.html", {
        "orders": orders,
        "orders_until_discount": orders_until_discount,
    })


@login_required
def order_list(request):
    orders = Order.objects.filter(user=request.user).order_by('id')
    for i, order in enumerate(orders, start=1):
        order.user_number = i
    return render(request, "main/order_list.html", {"orders": orders})


@login_required
def create_order(request):
    if request.method == "POST":
        service_type = ', '.join(request.POST.getlist("service_type"))
        item_name = request.POST.get("item_name")
        branch = request.POST.get("branch")
        defect_description = request.POST.get("defect_description", "")
        desired_date = request.POST.get("desired_date")
        urgent = request.POST.get("urgent") == "True"

        total_price = request.POST.get("total_price", "0")
        price = int(total_price) if str(total_price).isdigit() else 0

        profile = request.user.profile

        if not branch or not item_name or not service_type or not desired_date:
            messages.error(request, "Заповніть всі обов'язкові поля")
            return redirect("catalog")

        Order.objects.create(
            user=request.user,
            service_type=service_type,
            item_name=item_name,
            branch=branch,
            defect_description=defect_description,
            desired_date=desired_date,
            urgent=urgent,
            total_price=price
        )
        return redirect("profile")

    return redirect("catalog")


@require_POST
@login_required
def pay_order(request, order_id):
    order = get_object_or_404(Order, id=order_id, user=request.user)
    order.is_paid = True
    order.save()

    profile = request.user.profile
    profile.paid_orders_count += 1
    if profile.paid_orders_count >= 3:
        profile.loyalty_status = "Постійний клієнт"
    profile.save()

    return JsonResponse({'status': 'ok'})


@require_POST
@login_required
def delete_order(request, order_id):
    order = get_object_or_404(Order, id=order_id, user=request.user)
    order.delete()
    return JsonResponse({'status': 'deleted'})


def order_detail(request, order_id):
    order = get_object_or_404(Order, id=order_id, user=request.user)
    return render(request, "main/order_detail.html", {"order": order})


def register(request):
    errors = {}

    if request.method == "POST":
        username = request.POST.get("username", "").strip()
        email = request.POST.get("email", "").strip()
        phone = request.POST.get("phone", "").strip()
        password = request.POST.get("password", "").strip()

        if len(username.split()) < 3:
            errors["username_error"] = "Введіть ПІБ у форматі: Імʼя Прізвище По-батькові"

        if not re.fullmatch(r'\+380\d{9}', phone):
            errors["phone_error"] = "Телефон має бути у форматі +380XXXXXXXXX"

        if len(password) < 8:
            errors["password_error"] = "Пароль має містити щонайменше 8 символів"

        if not errors.get("username_error") and User.objects.filter(username=username).exists():
            errors["username_error"] = "Користувач з таким ПІБ вже існує"

        if not errors:
            user = User.objects.create_user(username=username, email=email, password=password)
            profile, created = Profile.objects.get_or_create(user=user, defaults={'phone': phone})
            if not created and not profile.phone:
                profile.phone = phone
                profile.save()
            login(request, user)
            return redirect("home")

        return render(request, "main/register.html", {
            "errors": errors,
            "username": username,
            "email": email,
            "phone": phone,
        })

    return render(request, "main/register.html")


def login_view(request):
    error = None
    if request.method == "POST":
        username = request.POST.get("username")
        password = request.POST.get("password")
        user = authenticate(request, username=username, password=password)
        if user is not None:
            login(request, user)
            return redirect("home")
        else:
            error = "Неправильний логін або пароль"
    return render(request, "main/login.html", {"error": error})


def logout_view(request):
    logout(request)
    return redirect("home")


@require_http_methods(["GET", "POST"])
def reviews(request):
    if request.method == "POST":
        if not request.user.is_authenticated:
            return JsonResponse({"error": "Unauthorized"}, status=401)

        try:
            data = json.loads(request.body)
            title = data.get("title", "").strip()
            text = data.get("text", "").strip()
            rating = int(data.get("rating", 5))

            if not title or not text:
                return JsonResponse({"error": "Поля не заповнені"}, status=400)

            if not (1 <= rating <= 5):
                return JsonResponse({"error": "Невірний рейтинг"}, status=400)

            Review.objects.create(
                author=request.user,
                title=title,
                text=text,
                rating=rating,
                is_company=False,
            )
            return JsonResponse({"success": True})

        except (json.JSONDecodeError, ValueError):
            return JsonResponse({"error": "Невірні дані"}, status=400)

    company_reviews = Review.objects.filter(is_company=True).order_by("-created_at")[:3]
    user_reviews = Review.objects.filter(is_company=False).order_by("-created_at")[:3]

    return render(request, "main/reviews.html", {
        "company_reviews": company_reviews,
        "user_reviews": user_reviews,
    })
