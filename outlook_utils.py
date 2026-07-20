"""
outlook_utils.py

Reply-all email automation for the FF1 5k Lysing pull (step 7).
Reuses the same generic Outlook COM functions from the last project
(open_outlook_app, get_default_folder, find_latest_sent_email_by_subject,
create_reply_all_draft, attach_inline_image) — these aren't project-
specific. Body text format matches the real email chain shown by the
user (see chat history for reference screenshots).
"""

import win32com.client
import config


def open_outlook_app():
    """Open or connect to Outlook through COM."""
    return win32com.client.Dispatch("Outlook.Application")


def get_default_folder(folder_number):
    """Returns an Outlook default folder by its folder number.
    5 = Sent Items, 6 = Inbox."""
    outlook = open_outlook_app()
    namespace = outlook.GetNamespace("MAPI")
    return namespace.GetDefaultFolder(folder_number)


def find_latest_sent_email_by_subject(subject_text=None):
    """Finds the most recent sent email whose subject contains the
    given text. Defaults to config.VIRAL_EMAIL_SUBJECT_SEARCH."""
    subject_text = subject_text or config.VIRAL_EMAIL_SUBJECT_SEARCH
    sent_items = get_default_folder(5)
    messages = sent_items.Items
    messages.Sort("[SentOn]", True)

    for message in messages:
        try:
            subject = message.Subject or ""
            if subject_text.lower() in subject.lower():
                return message
        except AttributeError:
            continue

    return None


def create_reply_all_draft(message, display=False):
    """Creates a Reply All draft from an Outlook message. If
    display=True, opens the draft window for user review."""
    reply = message.ReplyAll()
    if display:
        reply.Display()
    return reply


def attach_inline_image(draft, image_path, content_id):
    """Attaches an image to an Outlook draft with a Content-ID so it
    can be displayed inline in the body."""
    attachment = draft.Attachments.Add(image_path)
    property_accessor = attachment.PropertyAccessor
    property_accessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        content_id
    )


def format_update_summary(new_lot_dates, all_processed_ok=True):
    """Builds the summary line matching the real format:
    'Control chart and boxplot updates with [date range].
    This update has [N] new datapoint(s).'

    new_lot_dates: list of date/datetime objects for the lots just
    processed, used to build the date range text.
    """
    if not new_lot_dates:
        return "No new datapoints this update."

    sorted_dates = sorted(new_lot_dates)
    count = len(sorted_dates)

    def fmt(d):
        return f"{d.month}/{d.day}"

    if count == 1:
        date_text = fmt(sorted_dates[0])
    else:
        date_text = f"{fmt(sorted_dates[0])} - {fmt(sorted_dates[-1])}"

    plural = "datapoint" if count == 1 else "datapoints"
    return (f"Control chart and boxplot updates with {date_text} "
            f"datapoint{'s' if count > 1 else ''}.  This update has "
            f"{count} new {plural}.")


def build_email_body(summary_line, xbar_content_id, boxplot_content_id):
    """Builds the HTML body with both charts inline, in the real
    order confirmed from the actual email chain: Xbar first, then
    Boxplot below it."""
    return f"""
    <div style="font-family: Calibri; font-size: 11pt;">
        <p>{summary_line}</p>
        <p><img src="cid:{xbar_content_id}" style="width:600px;"></p>
        <p><img src="cid:{boxplot_content_id}" style="width:600px;"></p>
    </div>
    """


def send_update_reply(new_lot_dates, chart_paths, display=True):
    """Full orchestration: find the latest sent email in the chain,
    create a reply-all draft, insert the summary text and both charts
    inline, and either send or just display for manual review.

    chart_paths: dict with "xbar" and "boxplot" keys (paths to PNG
    files, e.g. from minitab_utils.export_boxplot_and_xbar()).

    display=True (default): opens the draft for manual review/send,
    does NOT send automatically — matches the caution used throughout
    this project for anything QA-related. Set display=False and call
    draft.Send() yourself only once this has been validated.
    """
    latest_email = find_latest_sent_email_by_subject()
    if latest_email is None:
        raise ValueError(
            f"No sent email found matching "
            f"{config.VIRAL_EMAIL_SUBJECT_SEARCH!r} to reply to."
        )

    draft = create_reply_all_draft(latest_email, display=False)

    summary_line = format_update_summary(new_lot_dates)

    xbar_cid = "xbar_chart"
    boxplot_cid = "boxplot_chart"
    attach_inline_image(draft, chart_paths["xbar"], xbar_cid)
    attach_inline_image(draft, chart_paths["boxplot"], boxplot_cid)

    body_html = build_email_body(summary_line, xbar_cid, boxplot_cid)
    draft.HTMLBody = body_html + draft.HTMLBody

    if display:
        draft.Display()

    return draft