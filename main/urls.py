from django.urls import path
from . import views
from .views import HomePageView, ProposalsView, save_category_config

urlpatterns = [
    path('', HomePageView.as_view(), name='home'),

    path('register/', views.register,    name='register'),
    path('login/',    views.login_view,  name='login'),
    path('logout/',   views.logout_view, name='logout'),

    path('profile/', views.profile,    name='profile'),
    path('catalog/', views.catalog,    name='catalog'),
    path('reviews/', views.reviews,    name='reviews'),
    path('orders/',  views.order_list, name='order_list'),

    path('proposals/', ProposalsView.as_view(template_name='main/proposals.html'), name='proposals'),

    path('manager/',             views.manager_dashboard,  name='manager'),
    path('manager/export/',      views.export_orders_excel, name='export_excel'),
    path('catalog/save-config/',   views.save_category_config,   name='save_category_config'),
    path('catalog/delete-config/', views.delete_category_config, name='delete_category_config'),

    path('order/create/',                  views.create_order,  name='create_order'),
    path('order/<int:order_id>/',          views.order_detail,  name='order_detail'),
    path('order/<int:order_id>/edit/',     views.edit_order,    name='edit_order'),
    path('order/<int:order_id>/pay/',      views.pay_order,     name='pay_order'),
    path('order/<int:order_id>/delete/',   views.delete_order,  name='delete_order'),
]