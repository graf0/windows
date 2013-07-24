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

require 'chef/mixin/shell_out'
include Chef::Mixin::ShellOut

action :create do
  if @current_resource.exists
    Chef::Log.info "#{@new_resource} task already exists - nothing to do"
  else
    use_force = @new_resource.force ? '/F' : ''
    cmd =  "schtasks /Create #{use_force} /TN \"#{@new_resource.name}\" "
    cmd += "/SC #{@new_resource.frequency} "
    cmd += "/MO #{@new_resource.frequency_modifier} " if [:minute, :hourly, :daily, :weekly, :monthly].include?(@new_resource.frequency)
    cmd += "/TR \"#{@new_resource.command}\" "
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
    cmd += "/TR \"#{@new_resource.command}\" " if @new_resource.command
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

def load_current_resource
  @current_resource = Chef::Resource::WindowsTask.new(@new_resource.name)
  @current_resource.name(@new_resource.name)

  task_hash = load_task_hash(@current_resource.name)
  if task_hash[:TaskName] == '\\' + @new_resource.name
    @current_resource.exists = true
    if task_hash[:Status] == "Running"
      @current_resource.status = :running
    end
    @current_resource.cwd(task_hash[:Folder])
    @current_resource.command(task_hash[:TaskToRun])
    @current_resource.user(task_hash[:RunAsUser])
  end if task_hash.respond_to? :[]
end

private

def load_task_hash(task_name)
  Chef::Log.debug "looking for existing tasks"

  require "win32ole"
  task = {}
  service = WIN32OLE.new("Schedule.Service")
  service.Connect

  begin
    root_folder = service.GetFolder("\\")
    registered_task = root_folder.GetTask("\\#{task_name}")
    task_definition = registered_task.Definition
  rescue WIN32OLERuntimeError => e
    # no such task - ignore!
  else
    # set taskname - this way we now task exists
    task[:TaskName] = "\\#{task_name}"

    # get status
    task[:Status] = case registered_task.State
      when 0 then "Unknown" # TASK_STATE_UNKONWN
      when 1 then "Disabled" # TASK_STATE_DISABLED
      when 2 then "Queued" # TASK_STATE_QUEUED
      when 3 then "Ready" # TASK_STATE_READY
      when 4 then "Running" # TASK_STATE_RUNNING
      end

    # get actions
    task_definition.Actions.each do |action|
      if action.Type == 0 # TYPE_ACTION_EXEC
        task[:TaskToRun] = "#{action.Path} #{action.Arguments}".strip
        task[:Folder] = action.WorkingDirectory
        break
      end
    end

    # get user
    task[:RunAsUser] = task_definition.Principal.UserId
    task[:RunLevel] = case task_definition.Principal.RunLevel
      when 0 then :limited # TASK_RUNLEVEL_LUA
      when 1 then :highest # TASK_RUNLEVEL_HIGHEST
      end

    # # get triggers
    # task_definition.Triggers.each do |trigger|
    #   case trigger.Type
    #   when 1 # TASK_TRIGGER_TIME
    #     case trigger.Repetition.Interval
    #     when ""
    #       frequency, frequency_modifier = :once, nil
    #     when /^P.*T(\d+)H$/
    #       frequency, frequency_modifier = :hourly, $1.to_i
    #     when /^P.*T(\d+H)?(\d+M)$/
    #       hours =   ($1.to_s.chop || 0).to_i
    #       minutes =   ($2.to_s.chop || 0).to_i
    #       frequency, frequency_modifier = :minute, hours*60 + minutes
    #     end
    #   when 2 # TASK_TRIGGER_DAILY
    #     frequency, frequency_modifier = :daily, trigger.DaysInterval
    #   when 3 # TASK_TRIGGER_WEEKLY
    #     frequency, frequency_modifier = :weekly, trigger.WeeksInterval
    #   when 4 # TASK_TRIGGER_MONTHLY
    #     # count numer of bits set to 1 in MonthsOfYear property
    #     months_count = trigger.MonthsOfYear.to_s(2).split(//).inject(0) { |s,i| s + i.to_i }
    #     frequency, frequency_modifier = :monthly, months_count % 12
    #   when 6 # TASK_TRIGGER_IDLE
    #     frequency, frequency_modifier = :on_idle, nil
    #   when 8 # TASK_TRIGGER_BOOT
    #     frequency, frequency_modifier = :on_start, nil
    #   when 9 # TASK_TRIGGER_LOGON
    #     frequency, frequency_modifier = :on_logon, nil
    #   end

    #   task[:Frequency] = frequency
    #   task[:FrequencyModifier] = frequency_modifier

    #   break
    # end
  end

  task
end