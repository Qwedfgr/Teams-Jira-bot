# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.
import json
import os
import random
import shutil
from datetime import datetime as dt
from datetime import timedelta
import uuid

from botbuilder.core import ActivityHandler, TurnContext, CardFactory
from botbuilder.schema import ChannelAccount, Attachment, Activity, ActivityTypes, ConversationAccount

from docxtpl import DocxTemplate
from jira import JIRA
import yadisk
import dotenv

CARDS = [
    "resources/adaptive_card_example.json"
]

dotenv.load_dotenv()
JIRA_SERVER = os.environ['JIRA_SERVER']
JIRA_LOGIN = os.environ['JIRA_LOGIN']
JIRA_TOKEN = os.environ['JIRA_TOKEN']
YA_TOKEN = os.environ['YA_TOKEN']


class AdaptiveCardsBot(ActivityHandler):
    def _create_reply(self, activity, text=None, text_format=None):
        return Activity(
            type=ActivityTypes.message,
            timestamp=dt.utcnow(),
            from_property=ChannelAccount(
                id=activity.recipient.id, name=activity.recipient.name
            ),
            recipient=ChannelAccount(
                id=activity.from_property.id, name=activity.from_property.name
            ),
            reply_to_id=activity.id,
            service_url=activity.service_url,
            channel_id=activity.channel_id,
            conversation=ConversationAccount(
                is_group=activity.conversation.is_group,
                id=activity.conversation.id,
                name=activity.conversation.name,
            ),
            text=text or "",
            text_format=text_format or None,
            locale=activity.locale,
        )

    async def on_members_added_activity(
            self, members_added: [ChannelAccount], turn_context: TurnContext
    ):
        for member in members_added:
            if member.id != turn_context.activity.recipient.id:
                await turn_context.send_activity(
                    f"Добрый день! "
                    f"Чтобы увидеть окно формирования листов учета времени напишите что-нибудь"
                )

    async def on_message_activity(self, turn_context: TurnContext):
        if turn_context.activity.value and not turn_context.activity.value.get('x'):
            return
        elif turn_context.activity.value and turn_context.activity.value.get('x'):
            await self.get_worklogs(turn_context.activity.value, turn_context)
        message = Activity(
            text="",
            type=ActivityTypes.message,
            attachments=[self._create_adaptive_card_attachment()],
        )

        await turn_context.send_activity(message)

    def _create_adaptive_card_attachment(self) -> Attachment:
        random_card_index = random.randint(0, len(CARDS) - 1)
        card_path = os.path.join(os.getcwd(), CARDS[random_card_index])
        with open(card_path, "rb") as in_file:
            card_data = json.load(in_file)

        return CardFactory.adaptive_card(card_data)

    async def get_worklogs(self, data, turn_context):
        if data and not data.get('x'):
            del turn_context.activity.value['x']
            return

        basic_auth = (JIRA_LOGIN, JIRA_TOKEN)
        options = {"server": JIRA_SERVER}
        jira = JIRA(options, basic_auth=basic_auth)
        user = data['Email'].split('@')[0]
        date_from = dt.strptime(data['DateFrom'], '%Y-%m-%d').date()
        date_to = dt.strptime(data['DateTo'], '%Y-%m-%d').date()
        issues = jira.search_issues(
            f'timespent > 0 and worklogdate >= {dt.strptime(data["DateFrom"], "%Y-%m-%d").date()} '
            f'and worklogdate <= {dt.strptime(data["DateTo"], "%Y-%m-%d").date()} and worklogAuthor = {user}',
            fields='worklog, summary, project', maxResults=-1)
        worklogs = []
        total = 0
        for single_date in date_range(date_from, date_to):
            date = single_date.strftime("%Y-%m-%d")
            content = {
                'date': format_date(date),
                'agreement_date': format_date(data.get('AgreementDate', '')),
                'add_agreement_date': format_date(data.get('AdditionalAgreementDate', '')),
                'agreement_number': data.get('AgreementNumber', ''),
                'add_agreement_number': data.get('AdditionalAgreementNumber', ''),
                'employee': data['Employee']
            }
            for issue in issues:
                if issue.fields.worklog.total >= issue.fields.worklog.maxResults:
                    issue_worklogs = jira.worklogs(issue)
                else:
                    issue_worklogs = issue.fields.worklog.worklogs
                for work in issue_worklogs:
                    if dt.fromisoformat(work.started[:-5]).date() == dt.strptime(date, '%Y-%m-%d').date() \
                            and work.author.displayName == content['employee']:
                        worklogs.append({'date': format_date(str(dt.fromisoformat(work.started[:-5]).date())),
                                         'duration': round(work.timeSpentSeconds / 3600, 2),
                                         'issue_descr': issue.fields.summary,
                                         'issue': issue})
                        total += work.timeSpentSeconds
        content.update({'worklogs': worklogs,
                        'total': round(total / 3600, 2)})
        if total:
            await self.get_word_doc(content)
        shutil.make_archive(content['employee'], 'zip', content['employee'])
        shutil.rmtree(content['employee'])
        filename = f"{content['employee']}.zip"
        await turn_context.send_activity(f'[скачать архив с яндекс диска]({upload_to_ya(filename)})')
        os.remove(filename)

    async def get_word_doc(self, content):
        doc = DocxTemplate("resources/template.docx")
        doc.render(content)
        if not os.path.exists(content['employee']):
            os.mkdir(content['employee'])
        doc.save(f"{content['employee']}/{content['employee']} {content['date']}.docx")


def date_range(start_date, end_date):
    for n in range(int((end_date - start_date).days + 1)):
        yield start_date + timedelta(n)


def format_date(date):
    if not date:
        return
    return dt.strptime(date, '%Y-%m-%d').strftime("%d.%m.%Y")


def upload_to_ya(filename):
    ya = yadisk.YaDisk(token=YA_TOKEN)
    ya_path = f'/{uuid.uuid1()} {filename}'
    ya.upload(filename, ya_path)
    ya.publish(ya_path)
    return ya.get_download_link(ya_path)
