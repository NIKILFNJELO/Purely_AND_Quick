from django.urls import path
from . import views
from .views import HomePageView, ProposalsView

urlpatterns = [
    path('', HomePageView.as_view(), name='home'),

    path('register/', views.register, name='register'),
    path('login/', views.login_view, name='login'),
    path('profile/', views.profile, name='profile'),
    path('logout/', views.logout_view, name='logout'),

    path('reviews/', views.reviews, name='reviews'),
    path('proposals/', ProposalsView.as_view(template_name='main/proposals.html'), name='proposals'),
    path('catalog/', views.catalog, name='catalog'),

    path('orders/', views.order_list, name='order_list'),
    path('create_order/', views.create_order, name='create_order'),
    path('order/<int:order_id>/', views.order_detail, name='order_detail'),
    path('order/<int:order_id>/pay/', views.pay_order, name='pay_order'),
    path('order/<int:order_id>/delete/', views.delete_order, name='delete_order'),
]