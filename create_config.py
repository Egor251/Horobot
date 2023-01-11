try:
    import configparser
except ImportError:
    import ConfigParser as configparser


def create_config(path):
    """
    Create a config file
    """
    config = configparser.ConfigParser()
    config.add_section("Settings")
    config.set("Settings", "email", "")
    config.set("Settings", "password", "")
    config.set("Settings", "SMTP_server", "smtp.yandex.ru")
    config.set("Settings", "SMTP_port", "465")
    config.set("Settings", "POP3_server", "pop.yandex.com")
    config.set("Settings", "POP3_port", "995")
    config.set("Settings", "admin_mail", "")
    config.set("Settings", "kachanova_mail", "")

    with open(path, "w") as config_file:
        config.write(config_file)