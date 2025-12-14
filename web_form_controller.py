# web_form_controller.py
import json
import web_socket_server


def fill_field(field_id: str, value: str) -> bool:
    """
    Send a message to browser to fill a form field.
    """
    payload = {"type": "fill", "field": field_id, "value": value}
    return web_socket_server.send(payload)


def submit_form() -> bool:
    """
    Send a message to browser to trigger form submission.
    """
    payload = {"type": "submit"}
    return web_socket_server.send(payload)
