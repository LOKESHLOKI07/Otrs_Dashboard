from django.core.management.base import BaseCommand
from mailer.cron import hello

class Command(BaseCommand):
    help = 'Send report emails'

    def handle(self, *args, **options):
        reports = hello().objects.all()
        for report in reports:
            report.send_email()
