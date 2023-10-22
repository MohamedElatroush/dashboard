from django.urls import include, path
# import routers
from rest_framework import routers
from .views import *
from rest_framework_simplejwt.views import (
    TokenObtainPairView,
    TokenRefreshView,
)

# define the router
router = routers.DefaultRouter()

# define the router path and viewset to be used
router.register(r'user', UserViewSet, basename='user')

urlpatterns = [
     path('', include(router.urls)),
     path('api-auth/', include('rest_framework.urls')),
     path('api/token/', TokenObtainPairView.as_view(), name='token_obtain_pair'),
     path('api/token/refresh/', TokenRefreshView.as_view(), name='token_refresh'),
]