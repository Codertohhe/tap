var userImageLogo = "";
var autoDB;
var tgType = "";

function uniqueArray(arr) {
	var a = [];
	for (var i = 0, l = arr.length; i < l; i++)
		if (a.indexOf(arr[i]) === -1 && arr[i] !== "") a.push(arr[i]);
	return a;
}

function startScreen() {
	$("#testgenerator-type").hide();
	tgType = "manual";
	$("#typeAuto").remove();
	//$("#typeAuto").hide();
	document.getElementById("startScreen").style.display = "none";
	document.getElementById("activityArea").style.display = "block";
	$treeview.jstree("close_all");

	setScrollbar();
	handleStructure("j1_1");

	window.testType = "testpaper";
	//setHeightFix();
	/*jQuery.confirm({
		title: 'Select category',
	    content: '',
	    buttons: {
	        'Worksheet': function () {
				document.getElementById('student-version-m').checked = false;
				document.getElementById('student-version').checked = true;
				window.testType = "worksheet";
				setHeightFix();
	        },
	        'Test Paper': function () {
				document.getElementById('student-version-m').checked = true;
				document.getElementById('student-version').checked = false;
				window.testType = "testpaper";
				setHeightFix();
	        }
	    }
	});*/

	var outcomes = getUniqueData("outcome");
	var outcomeHTML = "";
	for (var i = 0; i < outcomes.length; i++) {
		outcomeHTML +=
			'<option selected value="' +
			outcomes[i] +
			'">' +
			outcomes[i] +
			"</option>";
	}
	$("#manual-outcome").html(outcomeHTML);
	$("#manual-outcome").SumoSelect({
		captionFormatAllSelected: "All Selected",
		csvDispCount: 2,
	});
	$("#manual-difficulty").SumoSelect({
		captionFormatAllSelected: "All Selected",
		csvDispCount: 2,
	});

	$("#manual-outcome").on("sumo:closed", function (sumo) {
		//console.log(sumo);
	});

	$("#manual-difficulty").on("sumo:closed", function (sumo) {
		//console.log(sumo);
	});

	$("#btn-refresh-list").bind("click", function (e) {
		$(".jstree-clicked").trigger("click");
	});

	$("#totalQuestionsInTG").text(allQuestions().count());
}

function startAuto() {
	$("#testgenerator-type").hide();
	tgType = "auto";
	$("#activityArea").remove();
	//$("#activityArea").hide();
	document.getElementById("startScreen").style.display = "none";
	document.getElementById("typeAuto").style.display = "block";

	getChapters();

	var outcomes = getUniqueData("outcome");
	var outcomeHTML = "";
	for (var i = 0; i < outcomes.length; i++) {
		outcomeHTML +=
			'<li class="checkbox">' +
			'<label class="checkbox" for="outcome-auto-' +
			outcomes[i] +
			'">' +
			'<input id="outcome-auto-' +
			outcomes[i] +
			'" type="checkbox" value="' +
			outcomes[i] +
			'" checked="checked" class="input-mini"> ' +
			outcomes[i] +
			"</label>" +
			"</li>";
	}
	$("#auto-filter-outcome ul").html(outcomeHTML);

	var difficulty = getUniqueData("difficulty");
	difficulty = ["E", "M", "H"];
	var difficultyHTML = "";
	var difficultyObj = {
		E: "Easy",
		M: "Medium",
		H: "Hard",
	};
	//console.log(difficulty);
	for (var i = 0; i < difficulty.length; i++) {
		difficultyHTML +=
			'<li class="checkbox">' +
			'<label class="checkbox" for="difficulty-auto-' +
			difficulty[i] +
			'">' +
			'<input id="difficulty-auto-' +
			difficulty[i] +
			'" type="checkbox" value="' +
			difficulty[i] +
			'" checked="checked" class="input-mini"> ' +
			difficultyObj[difficulty[i]] +
			"</label>" +
			"</li>";
	}
	$("#auto-filter-difficulty ul").html(difficultyHTML);
	document.getElementById("chooseAll").checked = false;
	document.getElementById("select-all-qt-auto").checked = false;
	document.getElementById("select-all-qt-auto-tg").checked = false;

	//console.log(outcomes, difficulty);
	/*jQuery.confirm({
		title: 'Please select test bank type',
	    content: '',
	    buttons: {
	        'Worksheet': function () {
				window.modeEnableType = 'worksheet';
	        },
	        'Test Paper': function () {
				window.modeEnableType = 'testpaper';
	        }
	    }
	});*/
	setHeightFix();
}

function selectTestType() {
	$("#testgenerator-type").show();
	/*
	jQuery.confirm({
		title: 'Please select your choice from below',
	    content: '',
	    buttons: {
	        Auto: function () {
				startAuto();
	        },
	        Manual: function () {
				startScreen();
	        }
	    }
	});
	*/
}

function switchScreen(currentId) {
	$("#level1, #level2").hide();
	$("#" + currentId).show();

	//$('#questionsViewer :checkbox').prop('false');

	$("#explorer").empty();
	$("#showCartBtn").show();

	if (currentId == "level1") {
		jQuery(".jstree-clicked").trigger("click");
	}
	setHeightFix();
}

function closePreview() {
	$("#level3").hide();
	$("#level2").show();
	$("#userInfo").parent().children().show().end().end().hide();
}

function showPreview1() {
	//$('#userInfo').parent().children().hide().end().end().show();
	$("#userInfo").show();
	$("#pdfPreview").scrollTop(0);
	if (tgType == "auto") {
		if (
			window.tgduration != undefined &&
			window.tgduration != null &&
			window.tgduration != ""
		) {
			$("#time").val(window.tgduration);
		}
	}
}

function showPreview() {
	loadImage();
	$("#level2, #auto-level-2").hide();
	$("#level3, #auto-level-3").show();

	$("#previewHTML")
		.html($("#pdfPreview").html())
		.promise()
		.done(function () {
			var school = $("#pdfPreview #school").val();
			$("#previewHTML #school").replaceWith(
				'<span id="text-school">' + school + "</span>"
			);

			var subject = $("#pdfPreview #subject").val();
			$("#previewHTML #subject").replaceWith(
				'<span id="text-subject">' + subject + "</span>"
			);

			var class_t = $("#pdfPreview #class_t").val();
			$("#previewHTML #class_t").replaceWith(
				'<span id="text-class_t">' + class_t + "</span>"
			);

			var name = $("#pdfPreview #name").val();
			$("#previewHTML #name").replaceWith(
				'<span id="text-name">' + name + "</span>"
			);
			/*
		var roll = $('#pdfPreview #roll').val();
		$('#previewHTML #roll').replaceWith('<span id="text-roll">' + roll + '</span>');
		*/
			var time = $("#pdfPreview #time").val();
			$("#previewHTML #time").replaceWith(
				'<span id="text-time">' + time + "</span>"
			);
			/*
		var set = $('#pdfPreview #set').val();
		$('#previewHTML #set').replaceWith('<span id="text-set">' + set + '</span>');
		*/
			$("#previewHTML #showPrevBtn").remove();

			var totalMarks = 0;
			$("#previewHTML .marksInput").each(function (i, e) {
				//get mark
				var spcMark = $("#pdfPreview .marksInput").eq(i).val();
				totalMarks += parseInt(spcMark);

				var $marks = $('<div class="marksPreview"></div>').html(
					"[Marks:&nbsp;" + spcMark + "]"
				);
				$(this).replaceWith($marks);
			});

			$("<div></div>")
				.append(
					'<label>Total Marks:</label><span id="text-marks">' +
						totalMarks +
						"</div>"
				)
				.appendTo("#previewHTML #userInfo");
			$("#previewHTML #userInfo .clear")
				.detach()
				.appendTo("#previewHTML #userInfo");

			var $table = $("<table></table>");
			$table.append("<tr><td></td><td></td><tr>");
			$table.append("<tr><td></td><td></td><tr>");
			$table.append("<tr><td></td><td></td><tr>");

			var subjectobj =
				'<tr><td align="left" valign="top"><label class="topHeadLable">Subject:</label></td><td class="topHeadTd" align="left" valign="top"><span id="text-subject">' +
				subject +
				"</span></td></tr>";
			var classobj =
				'<tr><td align="left" valign="top"><label class="topHeadLable">Class:</label></td><td class="topHeadTd" align="left" valign="top"><span id="text-class_t">' +
				class_t +
				"</span></td></tr>";
			var nameobj =
				'<tr><td align="left" valign="top" width="100px"><label class="topHeadLable">Name:</label></td><td class="topHeadTd" align="left" valign="top"><span id="text-name">' +
				name +
				"</span></td></tr>";
			// var rollobj = '<tr><td align="left" valign="top"><label class="topHeadLable" style="width:73px";>Roll No:</label></td><td  class="topHeadTd" align="left" valign="top"><span id="text-roll">' + roll + '</span></td></tr>';

			var schoolobj =
				'<tr><td align="left" valign="top"><label class="topHeadLable">School:</label></td><td class="topHeadTd" align="left" valign="top"><span id="text-school">' +
				school +
				"</span></td></tr>";
			// var setobj = '<tr><td align="left" valign="top"><label class="topHeadLable" style="width:69px";>Set no:</label></td><td class="topHeadTd" align="left" valign="top"><span id="text-set">' + set + '</span></td></tr>';
			var timeobj =
				'<tr><td align="left" valign="top"><label class="topHeadLable">Time:</label></td><td class="topHeadTd" align="left" valign="top"><span id="text-time">' +
				convertMinsToHrsMins(time) +
				"</span></td></tr>";
			var marksobj =
				'<tr><td align="left" width="150px" valign="top"><label class="topHeadLable" style="width:118px";>Total Marks:</label></td><td class="topHeadTd" align="left" valign="top"><span id="text-marks">' +
				totalMarks +
				"</span></td></tr>";

			if (testType == "worksheet") {
				marksobj = "";
			}

			var tablestr =
				'<table width="100%">\
		<tr>\
			<td width="50%" align="left" valign="top"><table>' +
				schoolobj +
				classobj +
				nameobj +
				'</table></td>\
			<td width="50%" align="left" valign="top"><table>' +
				subjectobj +
				timeobj +
				marksobj +
				"</table></td>\
		</tr>\
	</table>";

			$table.appendTo("#userInfo");
			$("#previewHTML #userInfo").html(tablestr);

			var stdnt = document.getElementById("student-version");
			if (stdnt.checked == false) {
				//$('#previewHTML .question-answer').remove();
			}

			//$('#previewHTML h3').remove();
			$("#previewHTML").children().show();

			var num;
			$("#previewHTML img").each(function (i, e) {
				//var width = getImgWidth(this.src);
				//this.setAttribute('width', 600 + 'px');
				//$(this).parent().append('<span class="numbering">' + num + '</span> ');
			});

			$(
				"#previewHTML .mx-tools, #previewHTML .question-info, .summary-selection"
			).remove();
			$("#previewHTML").height("80%");

			$(".marksDiv span").remove();

			//$('#previewHTML').next().next().children().eq(0).prepend('1.');
			$("#previewHTML .numbering").each(function (i, e) {
				//num = $(this).text();
				//$(this).next().next().children().eq(0).prepend('<span class="numbering">' + num + '</span> ');
				//$(this).remove();
			});

			//$('#previewHTML').prepend(userImageLogo);

			// $("#previewHTML").prepend('<div style="opacity: 0;">&nbsp;</div>');

			setNumberingPreview();
		});
}

function getImgWidth(imagePath) {
	var img = document.createElement("img");
	img.setAttribute("src", imagePath);
	document.body.appendChild(img);
	var width = img.naturalWidth;
	return width;
	/*
	 var $img = $('<img />')
	 var width = $img.attr('src', imagePath).appendTo('body').width();
	 $img.remove();
	 return width;
	*/
}

function setScrollbar() {
	return;
	$("#collectionList").height($("#collectionList").height());
	$("#collectionList").width($("#collectionList").width());
	$("#collectionList").mCustomScrollbar("destroy");
	$("#collectionList").mCustomScrollbar({
		axis: "x", // vertical and horizontal scrollbar
		theme: "dark",
		scrollButtons: {
			enable: true,
		},
	});
}

function selectAllQuestion() {
	$("#questionsViewer input[type='checkbox']").not(":checked").trigger("click");
}

function getChapters() {
	$("#chapterList").empty();
	var chapters = allQuestions().distinct("chapter");
	for (var i = 0; i < chapters.length; i++) {
		$(function () {
			$("#chapterList").append(
				'<li class="checkbox">\
				<label class="checkbox" for="chapter' +
					i +
					'">\
					<input id="chapter' +
					i +
					'" type="checkbox" value=""  onClick="ifSelected(this)" class="input-mini">' +
					chapters[i] +
					"\
				</label>\
			</li>"
			);
		});
	}
	setHeightFix();
}

function selectQuestions() {
	var statusSelect = $("#chooseAll").prop("checked");
	if (statusSelect == true) {
		$("#chapterList .input-mini").prop("checked", "true");
	} else {
		$("#chapterList .input-mini").removeAttr("checked");
	}
}

function checkSelection() {
	var selectionCheck = $("#sidePanel-auto .input-mini:checked").length;
	if (selectionCheck == 0) {
		showMsgBox("Select at least one chapter from the left pane");
		return false;
	}

	var selectionOutcome = $("#auto-filter-outcome .input-mini:checked").length;
	if (selectionOutcome == 0) {
		showMsgBox("Select Source");
		return false;
	}

	var selectionDifficulty = $(
		"#auto-filter-difficulty .input-mini:checked"
	).length;
	if (selectionDifficulty == 0) {
		showMsgBox("Select at least one difficulty");
		return false;
	}

	$("#auto-page1").hide();
	$("#auto-page2").show();
	$("#sidePanel-auto").addClass("readonly-pannel");

	window.modeEnableType = "testpaper";

	loadWorksheet();
	//autoSelectMarks();
	$("#auto-reset-btn").show();
}

function autoSelectMarks() {
	$("#auto-page1").hide();
	$("#auto-page2").show();
	$("#sidePanel-auto").addClass("readonly-pannel");

	if (window.modeEnableType == "testpaper") {
		//console.log('hi');
		$("#auto-testpaper").prop("checked", true);
	} else {
		//console.log('bye');
		$("#auto-worksheet").prop("checked", false);
	}

	loadWorksheet();
}

function selectAutoAgain() {
	$("#auto-page2").hide();
	$("#auto-page1").show();
	$("#sidePanel-auto").removeClass("readonly-pannel");
	$("#auto-reset-btn").hide();
}

function showOptions(checkStatus, selectElement, selectAllEle) {
	if (checkStatus == true) {
		$(selectElement).removeAttr("disabled");
	} else {
		$(selectElement).attr("disabled", "disabled");
		$("#" + selectAllEle).prop("checked", false);
	}

	var worksheetType = document.querySelector(
		'input[name="worksheet-selection"]:checked'
	).value;
	if (worksheetType == "testpaper") {
		calulateTotalMarks();
	}
}

function ifSelected(ele) {
	if (ele.checked == false) {
		document.getElementById("chooseAll").checked = false;
	}
}

function loadWorksheet() {
	var worksheetType = document.querySelector(
		'input[name="worksheet-selection"]:checked'
	).value;
	var selectedList = "";

	worksheetType = "testpaper";

	$("#select-all-qt-auto, #select-all-qt-auto-tg").prop("checked", false);

	var selectedChapters = $("#chapterList :checked")
		.map(function (i, e) {
			return $.trim(e.parentNode.textContent);
		})
		.get();

	autoDB = new TAFFY();

	var selectedDifficutly = $("#auto-filter-difficulty :checked")
		.map(function (i, e) {
			return this.value;
		})
		.get();

	var selectedOutcome = $("#auto-filter-outcome :checked")
		.map(function (i, e) {
			return this.value;
		})
		.get();

	for (var i = 0; i < selectedChapters.length; i++) {
		allQuestions({ chapter: selectedChapters[i] }).each(function (r) {
			console.log(`question-${r.outcome}`);
			var outcomeIndex = selectedOutcome.indexOf(r.outcome);
			var difficultyIndex = selectedDifficutly.indexOf(r.difficulty);

			console.log(`outcomeIndex - ${outcomeIndex}`);
			console.log(`difficultyIndex - ${difficultyIndex}`);
			if (outcomeIndex != -1 && difficultyIndex != -1) autoDB.insert(r);
		});
		console.log(selectedChapters);
		console.log(autoDB);
	}

	if (autoDB().count() == 0) {
		//console.log('bad filter. no questions.')
	}

	var types = autoDB().distinct("type");

	console.log(`before shuffle - ${types}`);
	//var _temp_types = types.concat();
	types = shuffle_array(types);
	console.log(`after shuffle - ${types}`);

	$("#auto-worksheet-details, #auto-testpaper-details").empty();

	if (worksheetType == "worksheet") {
		$("#section-worksheet").show();
		$("#section-testpaper").hide();

		var worksheetHTML = "";
		for (var i = 0; i < types.length; i++) {
			var selectHTML = "";
			var total = autoDB({ type: types[i] }).count();
			var selectHTML = getSelect("value-edit-", total, i + 1);

			worksheetHTML +=
				'<tr>\
	<td>\
		<label class="checkbox" for="value-checkmark-' +
				(i + 1) +
				'">\
		<input id="value-checkmark-' +
				(i + 1) +
				'" type="checkbox" value="" class="input-option" onclick="showOptions(this.checked, \'#value-edit-' +
				(i + 1) +
				"', 'select-all-qt-auto')\">" +
				types[i] +
				'</label>\
	</td>\
	<td class="mxio">' +
				selectHTML +
				"</td>\
</tr>";
		}

		$("#auto-worksheet-details").append(worksheetHTML);
	} else {
		$("#section-worksheet").hide();
		$("#section-testpaper").show();

		var testpaperHTML = "";
		for (var i = 0; i < types.length; i++) {
			var selectHTML = "";
			var total = autoDB({ type: types[i] }).count();
			var selectHTML = getSelect("value-edit-", total, i + 1);
			//console.log(types[i]);
			testpaperHTML +=
				'<tr>\
	<td>\
		<label class="checkbox" for="value-checkmark-' +
				(i + 1) +
				'">\
		<input id="value-checkmark-' +
				(i + 1) +
				'" type="checkbox" value="" class="input-option" onclick="showOptions(this.checked, \'#value-input-' +
				(i + 1) +
				", #value-edit-" +
				(i + 1) +
				"', 'select-all-qt-auto-tg')\">" +
				types[i] +
				'</label>\
	</td>\
	<td class="mxio">' +
				selectHTML +
				'</td>\
	<td><font face="arial">x </font><input type="text" disabled id="value-input-' +
				(i + 1) +
				'" class="marksper" onkeyup="calculateMarks(' +
				(i + 1) +
				')" onkeypress="return isNumberKey(event)" /></td>\
	<td><font face="arial">= </font><span id="value-marks-' +
				(i + 1) +
				'">0</span></td>\
</tr>';
		}

		$("#auto-testpaper-details").append(testpaperHTML);

		$("#auto-duration").val("");
		$("#auto-maxmarks").val("");
		$("#total-marks-obj").val("");
	}
	setHeightFix();
}

function shuffle_array(types) {
	var _types = [];

	if (types.indexOf("Multiple choice questions") != -1)
		_types[0] = "Multiple choice questions";
	else _types[0] = "";

	if (types.indexOf("Fill in the blanks") != -1)
		_types[1] = "Fill in the blanks";
	else _types[1] = "";

	if (types.indexOf("True or False") != -1) _types[2] = "True or False";
	else _types[2] = "";

	if (types.indexOf("Match the Following") != -1)
		_types[3] = "Match the Following";
	else _types[3] = "";

	if (types.indexOf("Short answer type questions") != -1)
		_types[4] = "Short answer type questions";
	else _types[4] = "";

	if (types.indexOf("Long answer type questions") != -1)
		_types[5] = "Long answer type questions";
	else _types[5] = "";

	if (types.indexOf("Circle and underline") != -1)
		_types[6] = "Circle and underline";
	else _types[6] = "";

	if (types.indexOf("Activity and puzzle") != -1)
		_types[7] = "Activity and puzzle";
	else _types[7] = "";

	if (types.indexOf("Comprehension questions") != -1)
		_types[8] = "Comprehension questions";
	else _types[8] = "";

	if (types.indexOf("Write True or False") != -1)
		_types[9] = "Write True or False";
	else _types[9] = "";

	var apply_arr = _types.filter(function (el) {
		return el != "";
	});

	return (types = apply_arr);
}

function calculateMarks(lop) {
	var totalQuestion = $("#value-edit-" + lop).val();
	var marks = $("#value-input-" + lop).val();
	var total = totalQuestion * marks;
	//console.log(total);
	$("#value-marks-" + lop).text(total);
	calulateTotalMarks();
}

function calulateTotalMarks() {
	var total = 0;
	$("#auto-testpaper-details input[type='checkbox']:checked").each(function (
		i,
		e
	) {
		var $parent = $(e).parents("tr");

		var totalQuestions = $parent.find("select").val();
		var marks = $parent.find(".marksper").val();

		$parent.find("span").text(totalQuestions * marks);

		total += parseInt($(e).parents("tr").find("span").text());
	});

	document.getElementById("auto-maxmarks").value = total;
	document.getElementById("total-marks-obj").textContent = total;
}

function getSelect(eleId, total, idIndex) {
	if (total == 0) {
		return "";
	}

	var selectHTML =
		"<label><select onChange='calulateTotalMarks()' id='" +
		eleId +
		idIndex +
		"' disabled>";
	for (var i = 1; i <= total; i++) {
		selectHTML += "<option value='" + i + "'>" + i + "</option>";
	}
	selectHTML += "</select></label>";

	return selectHTML;
}

function submitAutoTGWorksheet() {
	var totalSelection = $(
		"#auto-worksheet-details .input-option:checked"
	).length;

	if (totalSelection == 0) {
		showMsgBox("Please select Question Type");
		return false;
	}

	jQuery.confirm({
		title: "Are you sure you want to continue?",
		content: "",
		buttons: {
			Yes: function () {
				submitAutoTGWorksheetV();
			},
			No: function () {},
		},
	});
}

function submitAutoTGWorksheetV() {
	var totalSelection = $(
		"#auto-worksheet-details .input-option:checked"
	).length;

	if (totalSelection == 0) {
		showMsgBox("Please select Question Type");
		return false;
	}

	var autoTGList = $("#auto-worksheet-details .input-option:checked")
		.map(function (i, e) {
			var obj = {};
			obj["title"] = $.trim($(this).parent().text());
			obj["total"] = $(this).parents("tr").find("select").val();
			return obj;
		})
		.get();

	var selectedAutoIds = [];
	for (var i = 0; i < autoTGList.length; i++) {
		var ids = [];
		autoDB({ type: autoTGList[i].title }).each(function (r) {
			ids.push(r.id);
		});

		//console.log(ids, autoTGList[i].total);
		var resultIds = randomList(ids, autoTGList[i].total);

		selectedAutoIds.push.apply(selectedAutoIds, resultIds);
	}

	$("#student-version").attr("checked", "checked");
	$("#student-version-m").removeAttr("checked");

	difficultySummary = {
		E: 0,
		M: 0,
		H: 0,
	};
	outcomeSummary = {};
	totalMarksSummary = 0;
	window.testType = "worksheet";
	addToTestAuto(selectedAutoIds);
}

function submitAutoTGTestpaper() {
	var duration = $("#auto-duration").val();
	if (duration == "") {
		showMsgBox("Please enter duration.");
		return false;
	}

	var totalSelected = $(
		"#auto-testpaper-details input[type='checkbox']:checked"
	).length;
	console.log(totalSelected, " totalSelection");
	if (totalSelected == 0) {
		showMsgBox("Please select at least one Question Types.");
		return false;
	}

	var marksTracker = {};

	var flag = false;
	var autoTGList = [];
	$("#auto-testpaper-details input[type='checkbox']:checked").each(function (
		i,
		e
	) {
		var marks = $(e).parents("tr").find(".marksper").val();
		var title = $.trim($(e).parent().text());
		var total = $(e).parents("tr").find("select").val();

		allQuestions({ type: title }).each(function (r) {
			r.mark = marks;
		});
		marksTracker[title] = marks;
		if (marks == "") {
			flag = true;
			showMsgBox(
				"Please enter marks per question for all selected question type."
			);
			return false;
		}
		autoTGList.push({
			title: title,
			total: total,
			marks: marks,
		});
	});

	if (flag == true) {
		return false;
	}

	jQuery.confirm({
		title: "Are you sure you want to continue?",
		content: "",
		buttons: {
			Yes: function () {
				submitAutoTGTestpaperV();
			},
			No: function () {},
		},
	});
}

function submitAutoTGTestpaperV() {
	var duration = $("#auto-duration").val();
	if (duration == "") {
		showMsgBox("Please enter duration.");
		return false;
	}

	var totalSelected = $(
		"#auto-testpaper-details input[type='checkbox']:checked"
	).length;
	if (totalSelected == 0) {
		showMsgBox("Please select at least one Question Types.");
		return false;
	}

	var marksTracker = {};

	var flag = false;
	var autoTGList = [];
	$("#auto-testpaper-details input[type='checkbox']:checked").each(function (
		i,
		e
	) {
		var marks = $(e).parents("tr").find(".marksper").val();
		var title = $.trim($(e).parent().text());
		var total = $(e).parents("tr").find("select").val();

		console.log(marks, title, total, " marks");
		allQuestions({ type: title }).each(function (r) {
			r.mark = marks;
			//console.log(r);
		});
		marksTracker[title] = marks;
		if (marks == "") {
			flag = true;
			showMsgBox(
				"Please enter marks per question for all selected question type."
			);
			return false;
		}
		autoTGList.push({
			title: title,
			total: total,
			marks: marks,
		});
		//console.log(allQuestions, allQuestions.length, autoTGList.length)
	});

	if (flag == true) {
		return false;
	}

	//console.log(autoTGList, autoTGList.length, ' autoTGList')
	var selectedAutoIds = [];
	for (var i = 0; i < autoTGList.length; i++) {
		autoDB({ type: autoTGList[i].title }).update({
			mark: String(autoTGList[i].marks),
		});

		var ids = [];
		autoDB({ type: autoTGList[i].title }).each(function (r) {
			ids.push(r.id);
		});
		//console.log( autoTGList[i].total)
		//console.log(ids, autoTGList[i].total);
		var resultIds = randomList(ids, autoTGList[i].total);

		selectedAutoIds.push.apply(selectedAutoIds, resultIds);

		console.log(selectedAutoIds, resultIds, " selectedAutoIds");
	}

	$("#student-version")[0].checked = false;
	$("#student-version-m")[0].checked = true;
	window.tgduration = duration;
	window.testType = "testpaper";

	addToTestAuto(selectedAutoIds);
}

function randomList(arr, len) {
	var randomQuestion = shuffleArray(arr);
	var resultArr = randomQuestion.slice(0, len);
	return resultArr;
	//return arr;
}

function shuffleArray(array) {
	for (var i = array.length - 1; i > 0; i--) {
		var j = Math.floor(Math.random() * (i + 1));
		var temp = array[i];
		array[i] = array[j];
		array[j] = temp;
	}
	return array;
}

function showMsgBox(msg) {
	jQuery.confirm({
		title: msg,
		content: "",
		buttons: {
			Ok: function () {},
		},
	});
}

function isNumberKey(evt) {
	var charCode = evt.which ? evt.which : evt.keyCode;
	if (charCode > 31 && (charCode < 48 || charCode > 57)) return false;
	return true;
}

function selectAll(status, selector) {
	if (status == true) {
		$(selector).find('input[type="checkbox"]').not(":checked").trigger("click");
	} else {
		$(selector).find('input[type="checkbox"]:checked').trigger("click");
	}
}

function convertMinsToHrsMins(minutes) {
	var result = "";
	var h = Math.floor(minutes / 60);
	var m = minutes % 60;
	//h = h < 10 ? '0' + h : h;
	//m = m < 10 ? '0' + m : m;
	if (h != 0) {
		result = h + " hour ";
	}
	if (m != 0) {
		result += m + " mins";
	}

	return $.trim(result);
}

function backToEdit() {
	$("#level3").hide();
	$("#level2").show();
	$("#level2 #userInfo").hide();
	$("#level2 #sContainIn").show();
	prepareSummary();
	$("#level3 #previewHTML").empty();

	$("#css-responsive").attr("href", "style/responsive1.css");
}

function backToEditAuto() {
	$("#auto-level-3").hide();
	$("#auto-level-2").show();
	$("#auto-level-2 #userInfo").hide();
	$("#auto-level-2 #sContainIn").show();
	prepareSummary();
	$("#auto-level-3 #previewHTML").empty();

	$("#css-responsive").attr("href", "style/responsive1.css");
}

function resetSummary() {
	difficultySummary = {
		E: 0,
		M: 0,
		H: 0,
	};
	outcomeSummary = {};
	totalMarksSummary = 0;
}

function setNumberingPreview() {
	$("#previewHTML .numbering").each(function (e) {
		var numberingHTML =
			"<span class='previewNumbering'>" + this.textContent + "</span>";
		var olddata = $(this).parent().find(".question-title").html();

		$(this)
			.parent()
			.find(".question-title")
			.html(numberingHTML + "<span class='aligntop'>" + olddata + "</span>");

		/* --- image top line fix --- */
		var $image = $(this).parents(".question-parent").find(".image");
		var $parentQ = $(this).parents(".question-parent");

		if ($image.length != 0) {
			$image.detach().appendTo($parentQ.find(".question-title"));
			$image.css("margin-left", "0%");
		}
		/* --- image top line fix --- */

		this.outerHTML = "";
	});
}

function setHeightFix() {
	var win = {
		height: $(window).height(),
		width: $(window).width(),
	};

	var explorerHeight = win.height - $("#level1 #topHeader").height();
	$("#level1 #level1_explorer").height(explorerHeight + "px");

	var qvh =
		explorerHeight -
		$("#level1 #stats").height() -
		$("#breadcrums").height() -
		20;
	$("#level1 #questionsViewer, #level1 #folders").height(qvh + "px");
	$("#level1 #questionsViewer, #level1 #folders").css("overflow", "auto");

	var mainframeH = win.height - $("#auto-level-1 .auto-toptitle").height();
	$("#auto-level-1 #mainFrame").height(mainframeH + "px");

	$("#auto-level-2 #toolbars").css("height", "auto");
	$("#level2 #toolbars").css("height", "auto");

	var toolbarHeight = $("#level2 #toolbars").height();
	if (toolbarHeight != 0) $("#level2 #toolbars").height(toolbarHeight + "px");

	var autoToolbarHeight = $("#auto-level-2 #toolbars").height();
	if (autoToolbarHeight != 0)
		$("#auto-level-2 #toolbars").height(autoToolbarHeight + "px");

	var top1Height = $("#level1 #topHeader").height();
	$("#level1 #topHeader .auto-tabs-section").css("height", top1Height + "px");

	$("#auto-level-1 .auto-tabs-section").css("height", "");
	var barheight = $("#auto-level-1 .auto-toptitle").height();
	$("#auto-level-1 .auto-tabs-section").css("height", barheight + "px");

	var heightTool = $("#toolbars").height();
	$("#toolbars .auto_level-buttons").css("height", heightTool + "px");

	if (win.width < 1000) {
		$(".book_name .tgTitle").css("font-size", "20px");
	}
}

$(document).ready(function (e) {
	/* enable breadcums */
	$("#explorer").on("click", ".breadLabel[data-link]", function (e) {
		var treeid = $(this).attr("data-link");
		$("#" + treeid + " > a").click();
	});
});

function confirmReload() {
	jQuery.confirm({
		title:
			"All your selection will be lost! Are you sure you want to go to home page?",
		content: "",
		buttons: {
			Confirm: function () {
				window.location.reload();
			},
			Cancel: function () {},
		},
	});
}

window.onbeforeunload = function () {
	return false;
};
