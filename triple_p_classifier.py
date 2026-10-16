import math
import re


OTHER_REVIEW = "OTHER - REVIEW"


def _normalize(text: str) -> str:
    return " ".join(re.findall(r"[A-Z0-9]+", text.upper()))


PEOPLE_KEYWORDS = tuple(
    _normalize(keyword)
    for keyword in (
        "BM", "BUSINESS MANAGER", "BUSINESS MANAGERS", "BUSINESS BANKER",
        "BUSINESS BANKERS", "PRIVATE BANKER", "BUSINESS CONSULTANT", "RM",
        "RELATIONSHIP MANAGER", "TELLERS", "CONSULTANT", "BANKER", "STAFF",
        "REPRESENTATIVE", "PERSONAL", "HUMAN", "MANAGEE", "BROKERS",
        "RELATIONSHIP", "COMMUNICATE", "TREATED", "INTERACTION", "PEOPLE",
        "CONTACTS", "CONTACT", "COMMUNICATION", "SOMEONE", "CUSTOMER SERVICE",
        "BM/RM", "PERSON", "WELCOME", "MANNER", "WALK", "HONESTY", "TRUST",
        "ATTENDING", "PROMISED", "UNDERSTAND", "CLIENTS",
    )
)

PROCESS_KEYWORDS = tuple(
    _normalize(keyword)
    for keyword in (
        "LETTER", "SECURE", "GUARANTEE", "ACCESS", "ACCESSIBLE", "ACTION",
        "ALLOWS", "APPROVED", "ASSIST", "ATTENDED", "ATTENTION", "CHANGE",
        "CONNECTIVITY", "CUSTOMER", "COMMUNICATION", "DEBIT", "DISRUPTING",
        "DOCUMENTATION", "EASIER", "EASILY", "ENQUIRIE", "FEEDBACK", "FINGER",
        "FOLLOWING", "FRAUD", "FUNCTIONALITY", "FUNDING", "PROVIDE",
        "INFORMATION", "ISSUES", "LENDING", "LONG", "NETWORK", "NOTIFICATION",
        "NOTIFIED", "OPENING", "ORDER", "OTP", "PERIOD", "PREVENT", "PROCESS",
        "QUERIES", "QUEUE", "QUEUES", "QUICKER", "QUICKLY", "REGULATION",
        "RESOLVE", "RETURN", "REVERSE", "QUERY", "RESOLUTION", "SECURITY",
        "SERVICE", "SETUP", "SPAM", "SPEED", "SYSTEM", "TIME", "TRANSACTION",
        "TRANSFER", "TURNAROUND", "UNAUTHORISED", "USER-FRIENDLY", "WAITING",
        "WORKING", "FASTER", "EXPIRE", "QUICK", "SANITIZER", "HELP", "LINK",
        "PROMISES", "CONVENIENT", "CLOSING", "EFFICIENCY", "URGENCY", "THROUGH",
        "SUNDAY", "SATURDAY", "DESKS", "INQUIRY", "POINTS",
        "VERIFYING MY ACCOUNT", "THE LAW ON CALL", "IF THERE IS A DISPUTE",
    )
)

PRODUCT_KEYWORDS = tuple(
    _normalize(keyword)
    for keyword in (
        "SHARIAH", "INVESTMENT", "SOLUTION", "EQUITIE", "CASH", "CHEQUES",
        "MONEY", "CARD", "FLEET", "ONLINE BANKING", "OVERDRAFT",
        "BUSINESS BANKING", "INVESTMENTS", "INSURANCE", "APP", "FNB BANKING APP",
        "CREDIT APPLICATION", "BANK CHARGES", "BANK FEES", "BANKING COSTS",
        "COSTS", "BUSINESS ACCOUNT", "CONTACT CENTRE", "CALL CENTRE",
        "CALL ACCOUNTS", "CALL ACCOUNT", "INVESTMENT ACCOUNTS", "PRICING",
        "SPEED POINT", "TRANSACTION", "FICA", "ONLINE", "MONEY MARKET", "FOREX",
        "CREDIT CARDS", "PRODUCTS", "FOREIGN", "E-BUCKS", "PRIVATE BANKING",
        "VEHICLE", "VAF", "CASH WITHDRAWAL", "E-WALLET", "BUSINESS", "PROPERTY",
        "FINANCE", "WESBANK", "SHARER", "TRADE", "BRANCHES", "FACILITIES",
        "CURRENCY", "UIF", "COVID", "MERCHANT", "ATM", "PAY", "PRODUCT",
        "EBACKS", "INTERNET", "BANKING", "LEVIES", "COST", "LOAN", "EXPENSIVE",
        "FEE", "E BUCKS", "NEW ACCOUNT", "FINANCING", "CREDIT", "LOANS",
        "INTEREST", "RATES", "CHARGE", "FEES", "TRAINING", "ADVICE", "INCONTACT",
        "UPDATING", "OPTIONS", "NOTIFICATIONS", "BENEFITS", "REWARDS", "EBUCKS",
        "ACCOUNTS", "STATEMENTS", "TRANSACTIONS", "NEW PRODUCTS", "PAYMENTS",
        "OPTION", "GIVING FUNDS FOR COMPANY ASSETS",
    )
)

NONE_KEYWORDS = tuple(
    _normalize(keyword)
    for keyword in (
        "NO IMPROVEMENT", "T THINK", "DON?T KNOW", "ALL FINE FROM MY SIDE",
        "N/A", "NOTHING", "NOTHING TO IMPROVE", "ANYTHING", "I AM SATISFIED",
        "I HAVE NO ISSUES", "NO COMMENT", "NONE", "HAPPY WITH EVERYTHING",
        "NO COMPLAINT AT THE MOMENT", "SATISFIED", "I AM HAPPY",
        "NO IMPROVEMENTS", "ALL GOOD", "ALL HAPPY", "ALL IN ORDER",
        "CANT THINK OF ANY", "DONT KNOW", "FUCK", "GOOD", "I DO NOT HAVE",
        "T HAVE", "HAVE ANSWERS", "ANY ISSUES", "NA", "NOT APPLICABLE",
        "NOT MUCH", "NOT NOW", "NOT SURE", "NOT AT THE MOMENT",
        "NOT SURE OF ANY", "NOT THAT I THINK OF", "-1", "EVERYTHING",
        "EXCELLENT", "KNOW", "ALL IS FINE", "ANY PROBLEM", "NO ASPECTS",
        "NO ASPECT FOR FNB TO IMPROVE ON",
        "NO SINGLE IMPORTANT ASPECT THAT FNB NEEDS TO IMPROVE ON",
        "NO ASPECT", "NO IMPORTATNT ASPECT", "NO IMPORTANT ASPECT",
        "NO SINGLE MOST IMPORTANT ASPECT", "NO IMPROVMENTS",
    )
)


def _contains_keyword(text: str, keywords: tuple[str, ...] | list[str]) -> bool:
    return any(f" {keyword} " in f" {text} " for keyword in keywords)


def classify_case_description(description: object) -> str:
    if description is None or (
        isinstance(description, float) and math.isnan(description)
    ):
        return OTHER_REVIEW

    description_text = str(description).strip()
    if description_text.casefold() in {"<na>", "nan", "none"}:
        return OTHER_REVIEW

    text = _normalize(description_text)
    if not text:
        return OTHER_REVIEW
    if _contains_keyword(text, NONE_KEYWORDS):
        return "NONE"

    product = _contains_keyword(text, PRODUCT_KEYWORDS)
    people = _contains_keyword(text, PEOPLE_KEYWORDS)
    process = _contains_keyword(text, PROCESS_KEYWORDS)

    if product and people and process:
        return "ALL"
    if product and process:
        return "PRODUCT & PROCESS"
    if product and people:
        return "PRODUCT & PEOPLE"
    if people and process:
        return "PEOPLE & PROCESS"
    if product:
        return "PRODUCT ONLY"
    if process:
        return "PROCESS ONLY"
    if people:
        return "PEOPLE ONLY"
    return OTHER_REVIEW
