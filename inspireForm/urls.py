
from django.conf.urls import include, url
from django.contrib import admin

urlpatterns = [
    url(r'^admin/', admin.site.urls),
    url(r'', include('FormFilter.urls')),
    url(r'^accounts/', include('django.contrib.auth.urls')),

]
