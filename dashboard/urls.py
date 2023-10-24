from django.urls import include, path
# import routers
from rest_framework import routers
from .views import *
from rest_framework_simplejwt.views import (
    TokenRefreshView,
)

# define the router
router = routers.DefaultRouter()

# define the router path and viewset to be used
router.register(r'user', UserViewSet, basename='user')
router.register(r'activity', ActivityViewSet, basename='activity')


urlpatterns = [
     path('', include(router.urls)),
     path('register/', UserRegistrationViewSet.as_view({'post': 'create_user'}), name='user-registration'),
     path('api-auth/', include('rest_framework.urls')),
     path('api/token/', MyTokenObtainPairView.as_view(), name='token_obtain_pair'),
     path('api/token/refresh/', TokenRefreshView.as_view(), name='token_refresh'),
]