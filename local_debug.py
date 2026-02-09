import re

def clean_subject(subject):
    """Removes Re:, Fwd: prefixes for soft matching."""
    if not subject: return ""
    s = subject.strip()
    while True:
        new_s = re.sub(r'^(Re:|Fwd:|FW:|Отв:)\s*', '', s, flags=re.IGNORECASE).strip()
        if new_s == s: break
        s = new_s
    return s

def test_clean():
    cases = [
        ("Re: Subject", "Subject"),
        ("Re: Re: Subject", "Subject"),
        ("Fwd: Re: Otv: Subject", "Subject"),
        ("Subject", "Subject"),
        ("   Re:   Subject  ", "Subject")
    ]
    
    failed = False
    for inp, exp in cases:
        got = clean_subject(inp)
        if got == exp:
            print(f"[PASS] '{inp}' -> '{got}'")
        else:
            print(f"[FAIL] '{inp}' -> '{got}' (Expected '{exp}')")
            failed = True
            
    if failed:
        print("Test FAILED")
        exit(1)
    else:
        print("Test PASSED")
        exit(0)

if __name__ == "__main__":
    test_clean()
