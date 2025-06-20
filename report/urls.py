from django.urls import path
from django.views.decorators.csrf import csrf_exempt
from rest_framework.decorators import authentication_classes, permission_classes, api_view
from rest_framework.permissions import IsAuthenticated

from core.jwt_authentication import JWTAuthentication
from . import views

urlpatterns = [
    path('xlsx/worker_vouchers/',
         api_view(['GET'])(
             authentication_classes([JWTAuthentication])(
                 permission_classes([IsAuthenticated])(
                     views.download_worker_vouchers_xlsx
                 )
             )
         ),
         name='download_worker_vouchers_xlsx'),
    path(
        "<str:report_name>/<str:report_format>/",
        views.report,
        name="report",
    ),
    path(
        "<str:report_name>/<str:report_format>/<str:alternate>/",
        views.report,
        name="report",
    ),
    path("reportbro/designer", views.reportbro_designer, name="reportbro_designer"),
    path(
        "reportbro/preview",
        csrf_exempt(views.reportbro_previewer),
        name="reportbro_previewer",
    ),
]
