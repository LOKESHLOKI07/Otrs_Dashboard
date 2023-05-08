from django.contrib.sites.models import Site

Site.objects.create(
    id=1,
    domain='yourdomain.com',
    name='Your Site Name',
)
