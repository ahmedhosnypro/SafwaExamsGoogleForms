let fst_year_fst_term_subjects = {
    study_year_id: 1,
    term_id: 1,
    subjects: [
        {
            name: "العقيدة",
            form_url: "https://docs.google.com/forms/d/1Gqqtw34JkaODRZHXKGmW_untLYWq8-lhU8_hM62PapA/edit?usp=forms_home&ths=true",
            safwa_id: 1,
        }
    ]
}

// variables for putting the questions and answers in the right position
let question_row = 1
let right_answer_index = 17

// main function to run
function main() {
    extractForms()
}

function extractForms() {
    fst_year_fst_term_subjects.subjects.forEach( subject => {
        question_row = 1
        extractFormQuestions(fst_year_fst_term_subjects, subject)
    })
}


//insert header
function insertHeader(sheet) {
    let column = 1
    //title
    sheet.getRange(question_row, column++).setValue("title")
    //points
    sheet.getRange(question_row, column++).setValue("points")
    //lesson or subject
    sheet.getRange(question_row, column++).setValue("lesson or subject")
    //path
    sheet.getRange(question_row, column++).setValue("path")
    //term
    sheet.getRange(question_row, column++).setValue("TERM")
    //subject
    sheet.getRange(question_row, column++).setValue("subject")
    //lesson
    sheet.getRange(question_row, column++).setValue("lesson")
    //question type
    sheet.getRange(question_row, column++).setValue("type")
    //description
    sheet.getRange(question_row, column++).setValue("description")
    //choices count
    sheet.getRange(question_row, column++).setValue("choices count")
    //coices
    sheet.getRange(question_row, column++).setValue("choice1")
    sheet.getRange(question_row, column++).setValue("choice2")
    sheet.getRange(question_row, column++).setValue("choice3")
    sheet.getRange(question_row, column++).setValue("choice4")
    sheet.getRange(question_row, column++).setValue("choice5")
    sheet.getRange(question_row, column++).setValue("choice6")
    //right answer
    sheet.getRange(question_row, column++).setValue("answer")
    question_row++
}

// Iterate over all questions
function extractFormQuestions(studyYearSubjects, subject) {
    let form = FormApp.openById(getFormId(subject.form_url))
    //create new spreadsheet with form name
    let ssNewUrl = SpreadsheetApp.create(form.getTitle()).getUrl()
    var sheet = SpreadsheetApp.openByUrl(ssNewUrl).getSheets()[0]

    insertHeader(sheet)
    form.getItems().forEach((item) => {
        switch (item.getType()) {
            case FormApp.ItemType.MULTIPLE_CHOICE:
                insertMul(sheet, studyYearSubjects, item.asMultipleChoiceItem())
                break
            case FormApp.ItemType.CHECKBOX:
                insertCheckBoxQuestion(sheet, studyYearSubjects, item.asCheckboxItem())
                break
        }
    })
    Logger.log(ssNewUrl)
}

function addRowBasicInfo(sheet, column, studyYearSubjects, title, points) {
    //title
    sheet.getRange(question_row, column++).setValue(title)
    //points
    sheet.getRange(question_row, column++).setValue(points)
    //lesson or subject
    sheet.getRange(question_row, column++).setValue("subject")
    //path
    sheet.getRange(question_row, column++).setValue(studyYearSubjects.study_year_id)
    //term
    sheet.getRange(question_row, column++).setValue(studyYearSubjects.term_id)
    //subject
    sheet.getRange(question_row, column++).setValue(fst_year_fst_term_subjects.subjects[0].name)
    //lesson
    sheet.getRange(question_row, column++).setValue("")
    return column;
}

function insertMul(sheet, studyYearSubjects, question) {
    let column = 1
    let type = "MULTIPLE_CHOICE"
    let points = question.getPoints()
    if (points === 0) {
        return
    }

    let title = question.getTitle()

    {
        if (title === "المذهب" ||
            title === "تأكدت من كتابة البريد الإلكتروني صحيحًا" ||
            title === "تأكدت من كتابة الاسم كاملًا كما هو في الأوراق الرسمية باللغة العربية" ||
            title === "تأكدت من كتابة الرقم الجامعي كما هو في الأوراق الرسمية" ||
            title === "دخولك الاختبار أكثر من مرة يعرضك لخسارة جميع علاماتك" ||
            title === "أتعهد بعدم فتح الكتاب أو أي مصدر آخر أثناء الاختبار، وعدم نشر أو مناقشة أسئلة الاختبار إلا بعد انتهاء الوقت المحدد  والله على ما أقول  شهيد." ||
            title === "النوع"
        ) {
            return
        }
    }

    title = formatQuestionTitle(title)
    let choices = question.getChoices()

    column = addRowBasicInfo(sheet, column, studyYearSubjects, title, points);

    if (choices.length === 2) {
        if ((choices[0].getValue().includes("صح") && choices[1].getValue().includes("خطأ")) ||
            (choices[1].getValue().includes("صح") && choices[0].getValue().includes("خطأ"))) {
            type = "binary"
            //question type
            sheet.getRange(question_row, column++).setValue(type)
            //description
            sheet.getRange(question_row, column++).setValue("")
            //choices count
            sheet.getRange(question_row, column++).setValue("")
            addBinaryChoice(sheet, choices);
            return
        }
    }


    //question type
    sheet.getRange(question_row, column++).setValue(type)
    //description
    sheet.getRange(question_row, column++).setValue("")
    //choices count
    sheet.getRange(question_row, column++).setValue(choices.length)
    addMultipleChoices(sheet, column, choices);
    question_row++
}

function insertCheckBoxQuestion(sheet, studyYearSubjects, question) {
    let coloumn = 1
    let title = formatQuestionTitle(question.getTitle())
    let points = question.getPoints()
    if (points === 0) {
        return
    }

    let type = "CHECKBOX"
    let choices = question.getChoices()

    coloumn = addRowBasicInfo(sheet, coloumn, studyYearSubjects, title, points);
    //question type
    sheet.getRange(question_row, coloumn++).setValue(type)
    //description
    sheet.getRange(question_row, coloumn++).setValue("")
    //choices count
    sheet.getRange(question_row, coloumn++).setValue(choices.length)

    addCheckBoxChoice(sheet, coloumn, choices);

    question_row++
}

function addBinaryChoice(sheet, choices) {
    choices.forEach(function (choice) {
        if (choice.isCorrectAnswer()) {
            if (choice.getValue().includes("صح")) {
                sheet.getRange(question_row, right_answer_index).setValue("صح")
            } else {
                sheet.getRange(question_row, right_answer_index).setValue("خطأ")
            }
        }
    })
    question_row++
}

function addCheckBoxChoice(sheet, coloumn, choices) {
    choices.forEach(function (choice, index) {
        sheet.getRange(question_row, coloumn++).setValue(choice.getValue())
        if (choice.isCorrectAnswer()) {
            let correct_answer = sheet.getRange(question_row, right_answer_index).getValue()
            sheet.getRange(question_row, right_answer_index).setValue(correct_answer + (index + 1) + ",")
        }
    })
}

function addMultipleChoices(sheet, column, choices) {
    choices.forEach(function (choice, index) {
        sheet.getRange(question_row, column++).setValue(choice.getValue())
        if (choice.isCorrectAnswer()) {
            sheet.getRange(question_row, right_answer_index).setValue(index + 1)
        }
    })
}

function formatQuestionTitle(title) {
    return title.replace(/\(\d+\)-|\(\d+\) -|\(\d+\)|\d+ -|\d+-|\d+/, '').trim()
}

function getFormId(formUrl) {
    return formUrl.split("/")[5]
}




