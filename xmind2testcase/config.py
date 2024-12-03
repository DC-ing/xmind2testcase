from enum import Enum


class Priority(Enum):
    Z = 1
    A = 2
    B = 3
    C = 4

    @classmethod
    def default_value(cls):
        return cls.B.value

    @classmethod
    def default_name(cls):
        return cls.B.name

    @classmethod
    def values(cls):
        return [member.value for member in cls]

    @classmethod
    def names(cls):
        return [member.name for member in cls]

    @staticmethod
    def get_priority(pri_num):
        for member in Priority:
            if member.value == pri_num:
                return member.name
        return Priority.default_name()


class TestResult(Enum):
    NON_EXECUTION = (0, "未执行")
    PASS = (1, "PASS")
    FAILED = (2, "NG")
    BLOCKED = (3, "阻塞")

    @property
    def val(self):
        return self.value[0]

    @property
    def desc(self):
        return self.value[1]

    @classmethod
    def default_val(cls):
        return cls.NON_EXECUTION.val

    @classmethod
    def default_name(cls):
        return cls.NON_EXECUTION.name

    @classmethod
    def default_desc(cls):
        return cls.NON_EXECUTION.desc

    @classmethod
    def values(cls):
        return [member.val for member in cls]

    @classmethod
    def descs(cls):
        return [member.desc for member in cls]

    @classmethod
    def names(cls):
        return [member.name for member in cls]

    @staticmethod
    def get_desc(result_num):
        for member in TestResult:
            if member.val == result_num:
                return member.desc
        return TestResult.default_desc()


excel_dropdown = {
    "优先级": Priority.names(),
    "测试结果": TestResult.descs(),
}
merge_header = [
    "一级功能模块",
    "二级功能模块",
    "三级功能模块",
]

excel_header = [
    "编号",
    "一级功能模块",
    "二级功能模块",
    "三级功能模块",
    "优先级",
    "用例标题",
    "前置条件",
    "操作步骤",
    "预期结果",
    "测试结果",
    "JIRA 号",
    "编写人",
    "执行人",
    "备注",
]
