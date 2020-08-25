import glob
import csv
import numpy as np
import pandas as pd

from bokeh.plotting import figure
from bokeh.models import Range1d, Legend
from bokeh.palettes import Set1, Set2



# Data Import
# -----------

def read_performance_report(filename, date_format='%d.%m.'):
	""" Reads metadata from a single Zoom 'Performance Report'
	
	Args:
		filename (str): Name of Zoom CSV file to read
		date_format (str): A string column of formatted dates is
			generated using this format string (see time.strftime())
		
	Returns: Single-row pandas DataFrame with columns:
		- 'topic': Meeting title (topic)
		- 'meeting_id': Zoom meeting ID
		- 'datetime': Date and time of meeting
		- 'duration': Meeting duration in minutes
		- 'registered': Number of registered attendees
		- 'attended': Number of people who attended
		- 'attendance_rate': Attendance rate
		- 'questions': Number of Q&A questions recorded
		- 'date_str': Formatted date string for plotting (see date_format)
	"""
	prdata = []
	with open(filename, 'r', encoding='utf-8') as pr:
		csvdata = csv.reader(pr, delimiter=',', quotechar='"')
		for line in csvdata:
			prdata.append(line)

	meta = {'topic':            [str(prdata[3][0])],
			'meeting_id':       [str(prdata[3][1].replace('-', ''))],
			'datetime':         [str(prdata[3][2])],
			'duration':         [int(prdata[3][3])],
			'registered':       [int(prdata[6][0])],
			'attended':         [int(prdata[6][1])],
			'attendance_rate':  [float(prdata[6][2]) / 100],
			'questions':        [int(prdata[8][0])]
		   }

	df = pd.DataFrame(meta)

	# Convert dates to pandas datetime, but keep a date string for easy plotting
	df.loc[:, 'datetime'] = pd.to_datetime(df.datetime, format='%b %d, %Y %I:%M %p')
	df.loc[:, 'date_str'] = df.datetime.dt.strftime(date_format)
	return df


def read_all_performance_reports(folder='.'):
	""" Reads and aggregates a folder of zoom 'Performance Report' files
	into a DataFrame for analysis. 

	Args:
		folder (str): Path to a folder containing CSV files

	Returns: pandas DataFrame, one row per report
	"""
	data = pd.DataFrame(columns=['topic', 'meeting_id', 'datetime', 'duration',
								 'registered', 'attended', 'attendance_rate', 
								 'questions', 'date_str'])
	perfs = glob.glob('{:s}/*Performance Report.csv'.format(folder))
	for perf_rep in perfs:
		data = data.append(read_performance_report(perf_rep), 
						   ignore_index=True, sort=False)
	
	data = data.sort_values('datetime').reset_index(drop=True)
	return data

	
def read_poll_report(filename):
	""" Reads poll questions and answers from a single zoom 'Poll Report'
	into a long-format DataFrame (one row per attendee*question)
	
	Args:
		filename (str): Name of Zoom CSV file to read
	
	Returns:
	"""
	polldata = []
	with open(filename, 'r', encoding='utf-8') as pr:
		for line in pr.readlines():
			polldata.append(line)
	
	# Extract meeting ID, then drop header to get the actual CSV data
	meeting_id = str(polldata[3].split(',')[1].replace('-', ''))
	polldata = polldata[6:]
	
	# Read question-answer pairs and aggregate counts
	poll = []

	csvdata = csv.reader(polldata, delimiter=',', quotechar='"')
	for line in csvdata:
		num_questions = len(line) - 5 # skipping number, name, email, date, and final empty field
		for q in list(range(0, num_questions, 2)):
				q_text = str(line[4+q])   # question
				a_text = str(line[4+q+1]) # answer

				qdata = {'meeting_id': meeting_id,
						 'row_no': int(line[0]),
						 'name': str(line[1]),
						 'email': str(line[2]),
						 'datetime': str(line[3]),
						 'question': q_text,
						 'answer': a_text}
				poll.append(qdata)

	cols = ['meeting_id', 'row_no', 'name', 'email', 'datetime', 'question', 'answer']
	df = pd.DataFrame(poll, columns=cols)
	# Note: date/time format is actually different here than in the CSV header! -.-
	df.loc[:, 'datetime'] = pd.to_datetime(df.datetime, format='%b %d, %Y %H:%M:%S')
	return df


def read_all_poll_reports(folder='.'):
	""" Reads and aggregates a folder of zoom 'Poll Report' files
	into a DataFrame for analysis

	Args:
		folder (str): Path to a folder containing CSV files

	Returns: pandas DataFrame, one row per report*question*answer
	"""
	cols = ['meeting_id', 'row_no', 'name', 'email', 'datetime', 'question', 'answer']
	data = pd.DataFrame(columns=cols)
	polls = glob.glob('{:s}/*Poll Report.csv'.format(folder))
	for poll_rep in polls:
		data = data.append(read_poll_report(poll_rep), ignore_index=True)
	
	data = data.sort_values(['question', 'answer']).reset_index(drop=True)
	return data


def read_poll_report_counts(filename):
	""" Reads poll questions and answers from a single zoom 'Poll Report'
	and aggregates answer counts per question
	
	Args:
		filename (str): Name of Zoom CSV file to read
	
	Returns: pandas DataFrame with columns: 
		- 'meeting_id': Zoom meeting ID
		- 'question': Poll question text
		- 'answer': Answer text for corresponding poll question
		- 'count': Number of responses for a given answer
		- 'prop': Proportion of responses for a given answer
	"""
	polldata = []
	with open(filename, 'r', encoding='utf-8') as pr:
		for line in pr.readlines():
			polldata.append(line)
	
	# Extract meeting ID, then drop header to get the actual CSV data
	meeting_id = str(polldata[3].split(',')[1].replace('-', ''))
	polldata = polldata[6:]
	
	# Read question-answer pairs and aggregate counts
	poll = {}
	N = {}

	csvdata = csv.reader(polldata, delimiter=',', quotechar='"')
	for line in csvdata:
		num_questions = len(line) - 5 # skipping number, name, email, date, and final empty field
		for q in list(range(0, num_questions, 2)):
				q_text = str(line[4+q])   # question
				a_text = str(line[4+q+1]) # answer

				if q_text not in poll.keys():
					poll[q_text] = {}
				if a_text not in poll[q_text].keys():
					poll[q_text][a_text] = 1
				else:    
					poll[q_text][a_text] += 1

				# Count participants for each question to accurately calculate answer proportions
				# Note: zoom seems to export in blocks of questions depending on the number of 
				# respondents, i.e. questions answered by the same participants are tacked on horizontally
				if q_text not in N.keys():
					N[q_text] = 1
				else:
					N[q_text] += 1
		
	# Convert dict of counts to a DataFrame
	polltable = []
	for q in poll.keys():
		for a in poll[q].keys():
			polltable.append([str(meeting_id), q, a, int(poll[q][a]), float(poll[q][a]) / float(N[q]), int(N[q])])
	cols = ['meeting_id', 'question', 'answer', 'count', 'prop', 'responses']
	
	return pd.DataFrame(polltable, columns=cols)


def read_all_poll_report_counts(folder='.'):
	""" Reads and aggregates a folder of zoom 'Poll Report' files
	into a DataFrame for analysis, aggregating response counts. 

	Args:
		folder (str): Path to a folder containing CSV files

	Returns: pandas DataFrame, one row per report*question*answer
	"""
	cols = ['meeting_id', 'question', 'answer', 'count', 'prop', 'responses']
	data = pd.DataFrame(columns=cols)
	polls = glob.glob('{:s}/*Poll Report.csv'.format(folder))
	for poll_rep in polls:
		data = data.append(read_poll_report_counts(poll_rep), ignore_index=True)
	
	data = data.sort_values(['question', 'answer']).reset_index(drop=True)
	return data



# Visualization
# -------------

def plot_attendance_bokeh(perf_report, questions=False, title='Meeting Attendance',
						  legend = ['Registered', 'Attended', 'Questions']):
	""" Plot attendance across meetings using the Bokeh library
	
	Args:
		perf_report: DataFrame from read_performance_report
		questions (bool): if True, also plot number of Q&A questions
		title (str): Plot title
		legend ([str]): Legend labels 

	Returns: Bokeh figure object
	"""
	columns = ['registered', 'attended']
	if questions:
		columns.append('questions')
	colors = Set1[6]
	
	fig = figure(x_range=perf_report.date_str, title=title, plot_width=700, plot_height=400)

	c = 0
	for ix, d in enumerate(columns):
		fig.line(perf_report.date_str, perf_report[d], line_width=2, color=colors[c], legend_label=legend[ix])
		fig.scatter(perf_report.date_str, perf_report[d], size=6, color=colors[c])
		c += 1

	fig.xaxis[0].axis_label = 'Meeting Date'
	fig.yaxis[0].axis_label = 'Count'
	fig.y_range = Range1d(0, perf_report.registered.max()+10)
	fig.legend.location = 'top_left'

	return fig


def plot_question_bokeh(polldata, question, prop=True, answer_sort=None):
	""" Plot response proportions for a specific poll question 
	as line plot using the Bokeh library

	Args:
		polldata: pandas DataFrame from read_poll_report_counts
		question (str): Text of poll question to plot
		prop (bool): if True, plot proportions instead of counts
		answer_sort ([str]): if specified, responses will be stacked in the
			order specified here, e.g. ['a', 'b', 'c']

	Returns: Bokeh figure object
	"""
	# Select question data, either proportions or counts
	if prop:
		data = polldata.loc[(polldata.loc[:, 'question'] == question), ['date_str', 'answer', 'prop']]
		data.loc[:, 'prop'] = data.loc[:, 'prop'] * 100 # convert to percent for display
		a = list(data.answer.unique())
		data = data.pivot_table(columns='answer', values='prop', index=['date_str']).reset_index()
	else:
		data = polldata.loc[(polldata.loc[:, 'question'] == question), ['date_str', 'answer', 'count']]
		a = list(data.answer.unique())
		data.loc[:, 'count'] = data.loc[:, 'count'].astype('int64')
		data = data.pivot_table(columns='answer', values='count', index=['date_str']).reset_index()  
	
	data.fillna(0, inplace=True) # Ensure all cells are valid

	colors = Set2[8]
	colors = colors[0:len(a)]

	# Allow specifying answer order for stacking
	if answer_sort is not None:
		a = answer_sort

	fig = figure(x_range=data['date_str'], title=question, plot_width=600, plot_height=400)
	for ix, ans in enumerate(a):
		fig.line(data.date_str, data.loc[:, ans], line_width=2, color=colors[ix], legend_label=a[ix])
		fig.scatter(data.date_str, data.loc[:, ans], size=6, color=colors[ix])

	return fig


def plot_question_stacked_bokeh(polldata, question, prop=True, answer_sort=None):
	""" Plot response proportions for a specific poll question 
	as stacked bar plot using the Bokeh library

	Args:
		polldata: pandas DataFrame from read_poll_report_counts
		question (str): Text of poll question to plot
		prop (bool): if True, plot proportions instead of counts
		answer_sort ([str]): if specified, responses will be stacked in the
			order specified here, e.g. ['a', 'b', 'c']

	Returns: Bokeh figure object
	"""
	# Select question data, either proportions or counts
	if prop:
		data = polldata.loc[(polldata.loc[:, 'question'] == question), ['date_str', 'answer', 'prop']]
		data.loc[:, 'prop'] = data.loc[:, 'prop'] * 100 # convert to percent for display
		a = list(data.answer.unique())
		data = data.pivot_table(columns='answer', values='prop', index=['date_str']).reset_index()
	else:
		data = polldata.loc[(polldata.loc[:, 'question'] == question), ['date_str', 'answer', 'count']]
		a = list(data.answer.unique())
		data.loc[:, 'count'] = data.loc[:, 'count'].astype('int64')
		data = data.pivot_table(columns='answer', values='count', index=['date_str']).reset_index()  
	
	data.fillna(0, inplace=True) # Ensure all cells are valid

	colors = Set2[8]
	colors = colors[0:len(a)]

	# Allow specifying answer order for stacking, check if all answers exist in data
	if answer_sort is not None:
		for var in a:
			if var not in answer_sort:
				err = 'All answers for a given question must be included in answer_sort:\nQuestion: {:s}\nAnswers: {:s}'
				raise ValueError(err.format(question, str(a)))
		a = answer_sort

	# Shorten title to max. 60 characters
	if len(question) > 60:
		question = question[0:61] + '...'

	fig = figure(x_range=data['date_str'], title=question, plot_width=700, plot_height=400, toolbar_location='left')
	bars = fig.vbar_stack(a, x='date_str', width=0.4, source=data, color=colors)

	# Make custom legend - only way to move it outside the plot area in Bokeh (for now)
	l_items = []
	for ix, ans in enumerate(a):
		if len(ans) > 35:
			ans = ans[0:36] + '...'
		l_items.append((str(ans), [bars[ix]]))
	legend = Legend(items=l_items, location="center")
	fig.add_layout(legend, 'right')

	return fig

