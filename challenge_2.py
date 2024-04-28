from robocorp import browser
from robocorp.tasks import task

@task
def shortest_path_challenge():
    browser.configure(
        slowmo=100,
    )
