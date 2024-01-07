from django.core.management.base import BaseCommand
from dashboard.jobs.jobs import schedule, holidays


class Command(BaseCommand):
    help = 'Run the updater as a test'

    def handle(self, *args, **options):
        schedule()
        holidays()
        self.stdout.write(self.style.SUCCESS('Updater test completed successfully.'))
