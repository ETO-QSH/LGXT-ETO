import sys
import time
import msvcrt

from prompt_toolkit import Application
from prompt_toolkit.key_binding import KeyBindings
from prompt_toolkit.layout import Layout
from prompt_toolkit.layout.containers import HSplit, Window
from prompt_toolkit.layout.controls import FormattedTextControl
from prompt_toolkit.styles import Style

from lgxt import *

class MenuSelector:
    def __init__(self, options):
        self.options = options
        self.selected_index = 0
        self.app = None

    def get_tokens(self):
        tokens = []
        tokens.append(('class:total', '•   请选择一个选项' + '\n'))
        tokens.append(('class:border', '+' + '-' * 20 + '+\n'))
        for i, option in enumerate(self.options):
            if i == self.selected_index:
                tokens.append(('class:border', '|'))
                tokens.append(('class:selected-star', '☆ '))
                tokens.append(('class:selected-option', option.ljust(15)))
                tokens.append(('class:border', ' |\n'))
            else:
                tokens.append(('class:border', '|  '))
                tokens.append(('class:unselected', option.ljust(15)))
                tokens.append(('class:border', ' |\n'))
        tokens.append(('class:border', '+' + '-' * 20 + '+\n'))
        return tokens

    def get_key_bindings(self):
        kb = KeyBindings()

        @kb.add('up')
        def _(event):
            self.selected_index = (self.selected_index - 1) % len(self.options)

        @kb.add('down')
        def _(event):
            self.selected_index = (self.selected_index + 1) % len(self.options)

        @kb.add('enter')
        def _(event):
            event.app.exit(result=self.options[self.selected_index])

        @kb.add('c-c')
        def _(event):
            event.app.exit(result=None)

        @kb.add('c-b')
        def _(event):
            event.app.exit(result="back")

        return kb

    def run(self):
        style = Style.from_dict({
            'total': '#ffffff bold',
            'selected-star': '#ff8f8f bold',
            'selected-option': '#ffffff bold',
            'unselected': '#888888',
            'border': '#ffffff',
        })

        layout = Layout(
            HSplit([
                Window(
                    FormattedTextControl(self.get_tokens),
                    height=len(self.options) + 3,
                    wrap_lines=True,
                ),
            ])
        )

        self.app = Application(
            layout=layout,
            key_bindings=self.get_key_bindings(),
            full_screen=True,
            style=style,
            mouse_support=True,
        )

        result = self.app.run()
        if result:
            print(f'你选择了: {result}')
            return result
        else:
            print('已退出')
            return None

def getpass_with_asterisk(prompt='请输入密码：'):
    print(prompt, end='', flush=True)
    password = ''
    while True:
        ch = msvcrt.getch().decode('utf-8', 'ignore')
        if ch == '\r':
            print('')
            break
        elif ch == '\b':
            if len(password) > 0:
                password = password[:-1]
                msvcrt.putch(b'\b')
                msvcrt.putch(b' ')
                msvcrt.putch(b'\b')
        else:
            password += ch
            msvcrt.putch(b'*')
    return password


if __name__ == '__main__':
    print("项目地址：https://github.com/ETO-QSH/LGXT-ETO")
    print(f"{'-'*20}\n快捷键：\nCtrl+C to exit\nCtrl+B to back")
    while True:
        print('-' * 20)
        ul = input("请输入账号：")
        pl = getpass_with_asterisk()
        login_state, login_result = login(ul, pl)
        if login_state:
            print("登录成功，正在加载书本信息~")
            courses_state, courses_result = get_my_courses()
            if courses_state:
                courses = {item["bookName"]: item["courseId"] for item in courses_result}
                options = list(courses.keys())

                while True:
                    selector = MenuSelector(options)
                    res0 = selector.run()
                    if res0 == "back":
                        break
                    elif res0 is None:
                        break
                    else:
                        course_id = courses[res0]
                        works_state, works_result = get_course_works(course_id)
                        if works_state:
                            works = {item['workName']: (item['workId'], item['times'], item['grade']) for item in works_result}
                            work_options = list(sorted(works.keys(), reverse=True))

                            while True:
                                selector_work = MenuSelector(work_options)
                                res1 = selector_work.run()
                                if res1 == "back":
                                    break
                                elif res1 is None:
                                    break
                                else:
                                    work_id = works[res1][0]
                                    print(f"编号：{works[res1][0]}  次数：{works[res1][1]}  分数：{works[res1][2]}")

                                    while True:
                                        options_action = ["提交成绩为满分", "下载题目及答案", "退出程序并注销"]
                                        selector_action = MenuSelector(options_action)
                                        res2 = selector_action.run()
                                        if res2 == "back":
                                            break
                                        elif res2 == "提交成绩为满分":
                                            s = input("会消耗一次答题机会，出现问题后果自负 Y/N: ")
                                            if s.lower() == 'y':
                                                success, result = submit_answer(work_id, 100)
                                                print(result)
                                        elif res2 == "下载题目及答案":
                                            s = input("请选择导出文件格式 pdf/doc: ").lower()
                                            if s in ["pdf", "doc"]:
                                                questions = collect_all_questions(work_id)
                                                if questions:
                                                    for question in questions.values():
                                                        save_question(question, "temp")
                                                    if s == "pdf":
                                                        save_questions_to_pdf(questions, "temp", f"{res0}-{res1}")
                                                    else:
                                                        save_questions_to_word(questions, "temp", f"{res0}-{res1}")
                                                else:
                                                    print("题目已关闭")
                                        else:
                                            print("5秒后退出程序。。。")
                                            time.sleep(5)
                                            sys.exit()
                                    if res2 != "back":
                                        break
                            if res1 != "back":
                                break
                        else:
                            print(works_result)
                            break
            else:
                print(courses_result)
                break
        else:
            print(login_result)
            print('-'*20)
