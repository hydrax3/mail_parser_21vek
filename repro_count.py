
import email.utils
from dataclasses import dataclass

@dataclass
class MockMsg:
    subject: str
    from_: str

def extract_email(sender_str):
    """
    Extracts purity email address from a 'Name <email>' string.
    Returns lowercase email or empty string.
    """
    if not sender_str:
        return ""
    # parseaddr returns (realname, email_address)
    name, addr = email.utils.parseaddr(sender_str)
    return addr.lower()

def test_extraction():
    assert extract_email("Bob <bob@ya.ru>") == "bob@ya.ru"
    assert extract_email("bob@ya.ru") == "bob@ya.ru"
    assert extract_email("Alice Smith <alice.smith+tag@example.com>") == "alice.smith+tag@example.com"
    assert extract_email("") == ""
    print("Extraction tests passed.")

def test_counting_logic():
    target_sender_str = "Client User <client@domain.com>"
    target_email = extract_email(target_sender_str)
    
    thread_msgs = [
        MockMsg("Subj", "Client User <client@domain.com>"),       # Match
        MockMsg("Subj", "client@domain.com"),                     # Match (loose format)
        MockMsg("Re: Subj", "Support Agent <support@ours.com>"),  # No match
        MockMsg("Re: Subj", "Client User <client@domain.com>")    # Match
    ]
    
    count = 0
    for msg in thread_msgs:
        msg_from = extract_email(msg.from_)
        if msg_from == target_email:
            count += 1
            
    print(f"Count: {count}")
    assert count == 3
    print("Counting logic tests passed.")

if __name__ == "__main__":
    test_extraction()
    test_counting_logic()
