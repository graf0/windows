#
# Author:: Paul Mooring (<paul@opscode.com>)
# Cookbook Name:: windows
# Provider:: task
#
# Copyright:: 2012, Opscode, Inc.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#

# Task Scheduler 2.0 support

require "win32ole"
require 'chef/mixin/shell_out'
include Chef::Mixin::ShellOut

action :create do
  if @current_resource.exists
    Chef::Log.info "#{@new_resource} task already exists - nothing to do"
  else
    use_force = @new_resource.force ? '/F' : ''
    cmd =  "schtasks /Create #{use_force} /TN \"#{@new_resource.name}\" "
    schedule  = @new_resource.frequency == :on_logon ? "ONLOGON" : @new_resource.frequency
    cmd += "/SC #{schedule} "
    cmd += "/MO #{@new_resource.frequency_modifier} " if [:minute, :hourly, :daily, :weekly, :monthly].include?(@new_resource.frequency)
    cmd += "/SD \"#{@new_resource.start_day}\" " unless @new_resource.start_day.nil?
    cmd += "/ST \"#{@new_resource.start_time}\" " unless @new_resource.start_time.nil?
    cmd += command_option
    if @new_resource.user && @new_resource.password
      cmd += "/RU \"#{@new_resource.user}\" /RP \"#{@new_resource.password}\" "
    elsif (@new_resource.user and !@new_resource.password) || (@new_resource.password and !@new_resource.user)
      Chef::Log.fatal "#{@new_resource.name}: Can't specify user or password without both!"
    end
    cmd += "/RL HIGHEST " if @new_resource.run_level == :highest
    shell_out!(cmd, {:returns => [0]})
    @new_resource.updated_by_last_action true
    Chef::Log.info "#{@new_resource} task created"
  end
end

action :run do
  if @current_resource.exists
    if @current_resource.status == :running
      Chef::Log.info "#{@new_resource} task is currently running, skipping run"
    else
      cmd = "schtasks /Run /TN \"#{@current_resource.name}\""
      shell_out!(cmd, {:returns => [0]})
      @new_resource.updated_by_last_action true
      Chef::Log.info "#{@new_resource} task ran"
    end
  else
    Chef::Log.debug "#{@new_resource} task doesn't exists - nothing to do"
  end
end

action :change do
  if @current_resource.exists
    cmd =  "schtasks /Change /TN \"#{@current_resource.name}\" "
    cmd += command_option if @new_resource.command
    if @new_resource.user && @new_resource.password
      cmd += "/RU \"#{@new_resource.user}\" /RP \"#{@new_resource.password}\" "
    elsif (@new_resource.user and !@new_resource.password) || (@new_resource.password and !@new_resource.user)
      Chef::Log.fatal "#{@new_resource.name}: Can't specify user or password without both!"
    end
    shell_out!(cmd, {:returns => [0]})
    @new_resource.updated_by_last_action true
    Chef::Log.info "Change #{@new_resource} task ran"
  else
    Chef::Log.debug "#{@new_resource} task doesn't exists - nothing to do"
  end
end

action :delete do
  if @current_resource.exists
    use_force = @new_resource.force ? '/F' : ''
    cmd = "schtasks /Delete #{use_force} /TN \"#{@current_resource.name}\""
    shell_out!(cmd, {:returns => [0]})
    @new_resource.updated_by_last_action true
    Chef::Log.info "#{@new_resource} task deleted"
  else
    Chef::Log.debug "#{@new_resource} task doesn't exists - nothing to do"
  end
end

# Index is equal to state returned from api
TASK_STATE = [
  "Unknown",
  "Disabled",
  "Queued",
  "Ready",
  "Running"
]

# Constants used by api
TYPE_ACTION_EXEC = 0
TASK_TRIGGER_TIME = 1
TASK_TRIGGER_DAILY = 2
TASK_TRIGGER_WEEKLY = 3
TASK_TRIGGER_MONTHLY = 4
TASK_TRIGGER_IDLE = 6
TASK_TRIGGER_BOOT = 8
TASK_TRIGGER_LOGON = 9
TASK_RUNLEVEL_LUA = 0
TASK_RUNLEVEL_HIGHEST = 1

# Load current task status using OLE API. We don't use output of schtasks command, because
# this output is internationalized, so it depends on installed language version of windows.
# Commands and parameters used to manage scheduled tasks are not internationalized, and from vista
# and up schtasks command exists on all systems, so we can use it.
def load_current_resource
  @current_resource = Chef::Resource::WindowsTask.new(@new_resource.name)
  @current_resource.name(@new_resource.name)

  task_scheduler = WIN32OLE.new("Schedule.Service")
  task_scheduler.Connect

  begin
    root_folder = task_scheduler.GetFolder("\\")
    registered_task = root_folder.GetTask("\\#{@current_resource.name}")
    task_definition = registered_task.Definition
  rescue WIN32OLERuntimeError => e
    # no such task - ignore!
  else
    @current_resource.exists = true
    @current_resource.status = TASK_STATE.index(registered_task.State)

    # scan actions for first exec
    task_definition.Actions.each do |action|
      if action.Type == TYPE_ACTION_EXEC
        @current_resource.command("#{action.Path} #{action.Arguments}".strip)
        @current_resource.cwd(action.WorkingDirectory)
        break
      end
    end

    # get user
    @current_resource.user(task_definition.Principal.UserId)
    @current_resource.run_level(case task_definition.Principal.RunLevel
      when TASK_RUNLEVEL_LUA      then :limited
      when TASK_RUNLEVEL_HIGHEST  then :highest
      end)

    # detect frequency and modifier
    task_definition.Triggers.each do |trigger|
      case trigger.Type
      when TASK_TRIGGER_TIME
        case trigger.Repetition.Interval
        when ""
          frequency, frequency_modifier = :once, nil
        when /^P.*T(\d+)H$/
          frequency, frequency_modifier = :hourly, $1.to_i
        when /^P.*T(\d+H)?(\d+M)$/
          hours =   ($1.to_s.chop || 0).to_i
          minutes =   ($2.to_s.chop || 0).to_i
          frequency, frequency_modifier = :minute, hours*60 + minutes
        end
      when TASK_TRIGGER_DAILY
        frequency, frequency_modifier = :daily, trigger.DaysInterval
      when TASK_TRIGGER_WEEKLY
        frequency, frequency_modifier = :weekly, trigger.WeeksInterval
      when TASK_TRIGGER_MONTHLY
        # count numer of bits set to 1 in MonthsOfYear property
        months_count = trigger.MonthsOfYear.to_s(2).split(//).inject(0) { |s,i| s + i.to_i }
        frequency, frequency_modifier = :monthly, months_count % 12
      when TASK_TRIGGER_IDLE
        frequency, frequency_modifier = :on_idle, nil
      when TASK_TRIGGER_BOOT
        frequency, frequency_modifier = :on_start, nil
      when TASK_TRIGGER_LOGON
        frequency, frequency_modifier = :on_logon, nil
      end

      @current_resource.frequency(frequency)
      @current_resource.frequency_modifier(frequency_modifier)

      # we consider only first trigger
      break
    end
  end
end

# Returns properly escaped and formatted /TR options for schtasks
def command_option
  command, parameters = split_command(@new_resource.command)
  "/TR " + (parameters.empty? ? "\"#{command}\"" : "\"\\\"#{command}\\\" #{parameters}\"")
end


# Tt splits command into application and parameters parts.
#
# Returns: array in form: [application_name, parameters]
def split_command(command)
  case command
  when /^['"](.+?)['"]\s+(.+)/ then [$1, $2] # command in quotation marks with params
  when /^(\S+)\s+(.+)/ then [$1, $2] # command without quotation marks with params
  else [command, ""] # in any other case - just command, without params
  end
end