clear
clc

%Reading config file
config = readlines('config.ini');
addpath("packages\");

header_size = Config.readConfigNumber(config, 'Scenario_Template', 'header_size');
idx_constant_road_radius = Config.readConfigNumber(config, 'Scenario_Template', 'idx_constant_road_radius');
idx_road_friction_coefficient = Config.readConfigNumber(config, 'Scenario_Template', 'idx_road_friction_coefficient');
idx_road_gradient = Config.readConfigNumber(config, 'Scenario_Template', 'idx_road_gradient');
idx_desired_vehicle_speed = Config.readConfigNumber(config, 'Scenario_Template', 'idx_desired_vehicle_speed');
idx_acceleration = Config.readConfigNumber(config, 'Scenario_Template', 'idx_acceleration');
idx_test_run_id = Config.readConfigNumber(config, 'Scenario_Template', 'idx_test_run_id');
%idx_steering_front_angle = Config.readConfigNumber(config, 'Scenario_Template', 'idx_steering_front_angle');
%idx_steering_front_slew_rate = Config.readConfigNumber(config, 'Scenario_Template', 'idx_steering_front_slew_rate');
%idx_steering_rear_angle = Config.readConfigNumber(config, 'Scenario_Template', 'idx_steering_rear_angle');
%idx_steering_rear_slew_rate = Config.readConfigNumber(config, 'Scenario_Template', 'idx_steering_rear_slew_rate');
idx_torque_front_axle = Config.readConfigNumber(config, 'Scenario_Template', 'idx_torque_front_axle');
idx_torque_rear_axle = Config.readConfigNumber(config, 'Scenario_Template', 'idx_torque_rear_axle');
idx_torque_slew_rate = Config.readConfigNumber(config, 'Scenario_Template', 'idx_torque_slew_rate');
%idx_ride_height_front_left = Config.readConfigNumber(config, 'Scenario_Template', 'idx_ride_height_front_left');
%idx_ride_height_front_right = Config.readConfigNumber(config, 'Scenario_Template', 'idx_ride_height_front_right');
%idx_ride_height_rear_left = Config.readConfigNumber(config, 'Scenario_Template', 'idx_ride_height_rear_left');
%idx_ride_height_rear_right = Config.readConfigNumber(config, 'Scenario_Template', 'idx_ride_height_rear_right');
%idx_ride_height_slew_rate = Config.readConfigNumber(config, 'Scenario_Template', 'idx_ride_height_slew_rate');
%idx_unintended_braking_torque = Config.readConfigNumber(config, 'Scenario_Template', 'idx_unintended_braking_torque');
idx_very_slow_steering = Config.readConfigNumber(config, 'Scenario_Template', 'idx_very_slow_steering');
idx_slow_steering = Config.readConfigNumber(config, 'Scenario_Template', 'idx_slow_steering');
idx_braking = Config.readConfigNumber(config, 'Scenario_Template', 'idx_braking');
idx_ftti = Config.readConfigNumber(config, 'Scenario_Template', 'idx_ftti');

%scenario_list_path = Config.readConfig(config, 'Scenario_List', 'ftti_path');
%scenario_list_path = "D:\Huawei_FUSA\03_FuSa\01_Simulation\Simulation_Scenario_List.xlsx";
scenario_list_path = "Simulation_Scenario_List_FTTI.xlsx";
sheet_name = Config.readConfig(config, 'Scenario_Template', 'sheet_name');
headers = split(Config.readConfig(config, 'Testrun_List', 'headers'), ',');

skip_sheet_generation = Config.readConfigNumber(config, 'Testrun_List', 'skip_sheet_generation');
testrun_list_path = Config.readConfig(config, 'Testrun_List', 'path');

vsm_testrun_path = Config.readConfig(config, 'VSM_Testrun', 'path');

%Creating VSM Testrun structure from template and clearing testruns
load('Template.vsd', 'Data', '-mat');
vsmTestRuns = Data;
while ~isempty(vsmTestRuns.Track)
    vsmTestRuns.Track(1) = [];
end

opts = detectImportOptions(scenario_list_path);
opts.VariableTypes{4} = 'char';
opts.DataRange = 'A3';

scenarioListSheet = readtable(scenario_list_path, opts);
%scenarioListSheet = readtable(scenario_list_path, opts, 'Sheet', sheet_name, 'Range', header_size);
scenarioListCells = table2cell(scenarioListSheet);

testRunIDs = scenarioListCells(:, idx_test_run_id);
vsmTable = cell(1, length(testRunIDs));

tic
for iScenario = 1 : length(testRunIDs)
    vsmTestRuns.Track(iScenario) = Data.Track(1); %Creating a new scenario from the template
    scenarioName = sprintf('Scenario_%s', testRunIDs{iScenario}); %Getting scenario name from testrun IDs
    if contains(upper(scenario_list_path), 'FTTI')
        scenarioName = strcat('FTTI_', scenarioName);
    elseif contains(upper(scenario_list_path), 'ACCEPTANCE')
        scenarioName = strcat('ACCEPTANCE_', scenarioName);
    end

    %Setting scenario name for the actual testrun
    vsmTestRuns.Track(iScenario).DisplayName = scenarioName;
    vsmTestRuns.Track(iScenario).ManeuverStruct.SetupName = scenarioName;
    vsmTestRuns.Track(iScenario).SetupName = scenarioName;  

    %Checking if scenario is time or distance based    
    desiredSpeed = scenarioListCells{iScenario, idx_desired_vehicle_speed};
    if isnumeric(desiredSpeed)
        if desiredSpeed == 0 %Scenario is time based
            vsmTestRuns.Track(iScenario).CycleBase = 1;
        else %Scenario is distance based
            vsmTestRuns.Track(iScenario).CycleBase = 0;
        end
    else
        error('Error: Scenario %s: Vehicle speed has to be numeric.', testRunIDs{iScenario})
    end

    vehicleAcceleration = scenarioListCells{iScenario, idx_acceleration};
    desiredSpeed = scenarioListCells{iScenario, idx_desired_vehicle_speed};
    if vsmTestRuns.Track(iScenario).CycleBase == 1 %Time based simulation
        endDistance = 10;
        stepSize = .5;
        distanceFaultInjection = 4.99;
        stepNum = floor(endDistance / stepSize);
        distance = linspace(0, stepNum * stepSize, stepNum + 1);
        if stepNum * stepSize ~= endDistance
            distance = [distance, endDistance];
        end
        iFaultInjection = -1;
        for iDistance = 1 : length(distance) - 1
            if distance(iDistance) < distanceFaultInjection && distanceFaultInjection <= distance(iDistance + 1)
            	iFaultInjection = iDistance + 1;
            end
        end
        if iFaultInjection == -1 || iFaultInjection == length(distance)
            error('Error: Scenario %s: Specified fault injection could not be used.', testRunIDs{iScenario})
        end
        if distanceFaultInjection ~= distance(iFaultInjection)
            distance = [distance(1:iFaultInjection-1), distanceFaultInjection, distance(iFaultInjection:end)]';
        end
        faultActive = [zeros(iFaultInjection, 1); ones(length(distance) - iFaultInjection, 1)];
        
        controlMode = 22 * ones(size(distance))'; %Speed control mode is 22
        vehicleSpeed = desiredSpeed * ones(size(distance));
    else %Distance based simulation,
        %The track length depends on the curve and the speed.
        %The driver model requires longer track distance for higher speeds and more curved roads
%         switch desiredSpeed
%             case 10
%                 distanceFaultInjection = 60;
%             case 20
%                 distanceFaultInjection = 65;
%             case 40
%                 distanceFaultInjection = 85;
%             case 70
%                 distanceFaultInjection = 105;
%             case 80
%                 distanceFaultInjection = 150;
%             case 100
%                 distanceFaultInjection = 240;
%             case 130
%                 distanceFaultInjection = 430;
%             otherwise
%                 warning("Unexpected vehicle speed %d km/h", desiredSpeed)
%         end

        initialSpeed = 3.6 * sqrt(power(desiredSpeed / 3.6, 2) - 2 * vehicleAcceleration * 50);
        if initialSpeed < 5 && vehicleAcceleration>0 
            initialSpeed = 5;
            desiredSpeed = initialSpeed;
        end
        
        distanceFaultInjection = ceil((20 + (desiredSpeed / 3.6 * 10))/10) * 10;
        
        if distanceFaultInjection<50 && vehicleAcceleration>0 
            distanceFaultInjection = 60;
        end
        
        if desiredSpeed < 30          
            endDistance = ceil((distanceFaultInjection + (30 / 3.6) * 10) / 50) * 50;
        else
            endDistance = ceil((distanceFaultInjection + (desiredSpeed / 3.6) * 10) / 50) * 50;
        end
        
        stepSize = 1;
        radiusText = Config.get_value(scenarioListCells{iScenario, idx_constant_road_radius});
        if isnumeric(radiusText)
            if radiusText > 0 && radiusText < 25
                stepSize = 5;
            end    
        end
        if ~isnan(vehicleAcceleration) && isnumeric(vehicleAcceleration) && vehicleAcceleration ~= 0 %Acceleration during scenario
            distanceFaultInjection = distanceFaultInjection - 0.01;
            distanceAcceleration = distanceFaultInjection - 50;
            
            stepNum = floor(endDistance / stepSize);
            distance = linspace(0, stepNum * stepSize, stepNum + 1);
            if stepNum * stepSize ~= endDistance
                distance = [distance, endDistance];
            end
            if distanceAcceleration < 0
                error('Error: Scenario %s: Specified speed is too low.', testRunIDs{iScenario})
            end
            
            iAcceleration = -1;
            for iDistance = 1 : length(distance) - 1
                if distance(iDistance) < distanceAcceleration && distanceAcceleration <= distance(iDistance + 1)
                    iAcceleration = iDistance + 1;
                end
            end
            if iAcceleration == -1 || iAcceleration == length(distance)
                error('Error: Scenario %s: Specified acceleration could not be used.', testRunIDs{iScenario})
            end
            if distanceAcceleration ~= distance(iAcceleration)
                distance = [distance(1:iAcceleration-1), distanceAcceleration, distance(iAcceleration:end)];
            end
            
            iFaultInjection = -1;
            for iDistance = 1 : length(distance) - 1
                if distance(iDistance) < distanceFaultInjection && distanceFaultInjection <= distance(iDistance + 1)
                    iFaultInjection = iDistance + 1;
                end
            end
            if iFaultInjection == -1 || iFaultInjection == length(distance)
                error('Error: Scenario %s: Specified fault injection could not be used.', testRunIDs{iScenario})
            end
            if distanceFaultInjection ~= distance(iFaultInjection)
                distance = [distance(1:iFaultInjection-1), distanceFaultInjection, distance(iFaultInjection:end)];
            end
            
            distance = distance';
            
            controlMode = [22 * ones(iAcceleration, 1); 22 * ones(length(distance) - iAcceleration, 1)]; %Speed control mode is 22 and acceleration mode is 24
            faultActive = [zeros(iFaultInjection, 1); ones(length(distance) - iFaultInjection, 1)];

            %initialSpeed = 3.6 * sqrt(power(desiredSpeed / 3.6, 2) - 2 * vehicleAcceleration * 50);
            %if initialSpeed < 5
            %    initialSpeed = 5;
            %end
            vehicleSpeed = [initialSpeed * ones(iAcceleration, 1); 0 * ones(length(distance) - iAcceleration, 1)];
            for index = iAcceleration+1 : length(distance)-1
                vehicleSpeed(index)=  sqrt(2*vehicleAcceleration*12960*(distance(index)-distance(index-1))*10^-3+vehicleSpeed(index-1)^2);
                if vehicleAcceleration>0 && vehicleSpeed(index)< 5
                    vehicleSpeed(index)= 5;
                end
                    
            end
            %vehicleSpeed = [initialSpeed * ones(iAcceleration, 1); vehicleAcceleration / 9.81 * ones(length(distance) - iAcceleration, 1)];

        else %Constant speed driving
            distanceFaultInjection = distanceFaultInjection - 0.01;
            stepNum = floor(endDistance / stepSize);
            distance = linspace(0, stepNum * stepSize, stepNum + 1);
            if stepNum * stepSize ~= endDistance
                distance = [distance, endDistance];
            end
            iFaultInjection = -1;
            for iDistance = 1 : length(distance) - 1
                if distance(iDistance) < distanceFaultInjection && distanceFaultInjection <= distance(iDistance + 1)
                    iFaultInjection = iDistance + 1;
                end
            end
            if iFaultInjection == -1 || iFaultInjection == length(distance)
                error('Error: Scenario %s: Specified fault injection could not be used.', testRunIDs{iScenario})
            end
            if distanceFaultInjection ~= distance(iFaultInjection)
                distance = [distance(1:iFaultInjection-1), distanceFaultInjection, distance(iFaultInjection:end)]';
            end
            faultActive = [zeros(iFaultInjection, 1); ones(length(distance) - iFaultInjection, 1)];

            controlMode = 22 * ones(size(distance))'; %Speed control mode is 22
            vehicleSpeed = desiredSpeed * ones(size(distance));
        end
    end
    
    vsmArray = zeros(length(distance), length(headers));
    vsmTable{iScenario} = array2table(vsmArray);
    vsmTable{iScenario}.Properties.VariableNames = headers;

    if vsmTestRuns.Track(iScenario).CycleBase == 1 %Time based simulation
        vsmTable{iScenario}.time = distance;
        vsmTable{iScenario}.distance = zeros(size(distance));
    else %Distance based simulation
        vsmTable{iScenario}.distance = distance;
        vsmTable{iScenario}.time = zeros(size(distance));
    end
    
    %Checking the validity of distance (or time) steps
    if length(distance) < 2
        error('Error: Scenario %s: At least 2 steps are needed for each scenario.', testRunIDs{iScenario})
    elseif ~isnumeric(distance)
        error('Error: Scenario %s: Scenario distance (or time) steps are not valid.', testRunIDs{iScenario})
    elseif ~issorted(distance)
        error('Error: Scenario %s: Scenario distance (or time) steps are not increasing.', testRunIDs{iScenario})
    end

    vsmTestRuns.Track(iScenario).AddDMD1Map.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).AddDMD1Map.AddDMD1 = ones(1, length([distance(1); distance(end)]));
    
    vsmTestRuns.Track(iScenario).AddDMD2Map.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).AddDMD2Map.AddDMD2 = zeros(1, length([distance(1); distance(end)]));
    
    vsmTestRuns.Track(iScenario).AddDMD3Map.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).AddDMD3Map.AddDMD3 = 1.2882297539194254E-231*(ones(1, length([distance(1); distance(end)]))); %vsmTestRuns.Track(iScenario).AddDMD3Map.AddDMD3(1)
    
    vsmTestRuns.Track(iScenario).AddDMD4Map.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).AddDMD4Map.AddDMD4 = -1*(ones(1, length([distance(1); distance(end)])));
    
    vsmTestRuns.Track(iScenario).AddDMD5Map.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).AddDMD5Map.AddDMD5 = zeros(1, length([distance(1); distance(end)]));
    
    vsmTestRuns.Track(iScenario).AddDMD6Map.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).AddDMD6Map.AddDMD6 = zeros(1, length([distance(1); distance(end)]));
    
    vsmTestRuns.Track(iScenario).AddDMD7Map.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).AddDMD7Map.AddDMD7 = zeros(1, length([distance(1); distance(end)]));
    
    %Setting brake pressures
    vsmTestRuns.Track(iScenario).AddpBrakeMap.Dist = [distance(1); distance(end)]; 
    vsmTestRuns.Track(iScenario).AddpBrakeMap.AddpBrake = zeros(1, length(vsmTestRuns.Track(iScenario).AddpBrakeMap.Dist));

    %Setting clutch state
    vsmTestRuns.Track(iScenario).ClutchMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).ClutchMap.Clutch = [0, 0];

    %Checking and setting control mode and vehicle speed or acceleration
    vsmTestRuns.Track(iScenario).ControlMode = controlMode;
    vsmTable{iScenario}.mode = repmat("", size(vsmTable{iScenario}.mode));
    for iMode = 1 : length(controlMode)
        switch controlMode(iMode)
            case 22
                vsmTable{iScenario}.vehicleSpeed(iMode) = vehicleSpeed(iMode);
                vsmTable{iScenario}.mode(iMode) = "Speed";
            case 24
                vsmTable{iScenario}.acceleration(iMode) = vehicleSpeed(iMode) * 9.80665;
                vsmTable{iScenario}.mode(iMode) = "Acceleration";
            otherwise
                error('Error: Scenario %s: Control mode value is invalid.', testRunIDs{iScenario})
        end
    end
    vsmTestRuns.Track(iScenario).v = vehicleSpeed;    
       
    %Setting customer channels (and dealing with NaN's)    
    for i = 1 : length(distance)
        vsmTestRuns.Track(iScenario).CustomerChannels(i,1) = faultActive(i);
        vsmTestRuns.Track(iScenario).CustomerChannels(i,2) = 0; %Config.get_number(scenarioListCells{iScenario, idx_steering_front_angle});
        vsmTestRuns.Track(iScenario).CustomerChannels(i,3) = 0; %Config.get_number(scenarioListCells{iScenario, idx_steering_front_slew_rate});
        vsmTestRuns.Track(iScenario).CustomerChannels(i,4) = 0; %Config.get_number(scenarioListCells{iScenario, idx_steering_rear_angle});
        vsmTestRuns.Track(iScenario).CustomerChannels(i,5) = 0; %Config.get_number(scenarioListCells{iScenario, idx_steering_rear_slew_rate});
        vsmTestRuns.Track(iScenario).CustomerChannels(i,6) = Config.get_value(scenarioListCells{iScenario, idx_torque_front_axle});
        vsmTestRuns.Track(iScenario).CustomerChannels(i,7) = Config.get_value(scenarioListCells{iScenario, idx_torque_rear_axle});
        vsmTestRuns.Track(iScenario).CustomerChannels(i,8) = Config.get_number(scenarioListCells{iScenario, idx_torque_slew_rate});
        vsmTestRuns.Track(iScenario).CustomerChannels(i,9) = 0; %Config.get_number(scenarioListCells{iScenario, idx_ride_height_front_left});
        vsmTestRuns.Track(iScenario).CustomerChannels(i,10) = 0; %Config.get_number(scenarioListCells{iScenario, idx_ride_height_front_right});
        vsmTestRuns.Track(iScenario).CustomerChannels(i,11) = 0; %Config.get_number(scenarioListCells{iScenario, idx_ride_height_rear_left});
        vsmTestRuns.Track(iScenario).CustomerChannels(i,12) = 0; %Config.get_number(scenarioListCells{iScenario, idx_ride_height_rear_right});
        vsmTestRuns.Track(iScenario).CustomerChannels(i,13) = 0; %Config.get_number(scenarioListCells{iScenario, idx_ride_height_slew_rate});
        vsmTestRuns.Track(iScenario).CustomerChannels(i,14) = NaN; %Config.get_number(scenarioListCells{iScenario, idx_unintended_braking_torque});
        vsmTestRuns.Track(iScenario).CustomerChannels(i,15) = Config.get_value(scenarioListCells{iScenario, idx_very_slow_steering});
        vsmTestRuns.Track(iScenario).CustomerChannels(i,16) = Config.get_value(scenarioListCells{iScenario, idx_slow_steering});
        vsmTestRuns.Track(iScenario).CustomerChannels(i,17) = Config.get_value(scenarioListCells{iScenario, idx_braking});
        ftti = Config.get_value(scenarioListCells{iScenario, idx_ftti});
        if ischar(ftti) && isempty(ftti)
            ftti = 1000;
        else
            ftti = ftti / 1000;
        end
        vsmTestRuns.Track(iScenario).CustomerChannels(i,18) = ftti;
        vsmTestRuns.Track(iScenario).CustomerChannels(i,19) = str2num(scenarioListCells{iScenario, idx_test_run_id});
    end

    vsmTable{iScenario}.faultActive = vsmTestRuns.Track(iScenario).CustomerChannels(:,1);
    vsmTable{iScenario}.Front_Steering_fault_deg = vsmTestRuns.Track(iScenario).CustomerChannels(:,2);
    vsmTable{iScenario}.Front_Steering_slew_rate_degps = vsmTestRuns.Track(iScenario).CustomerChannels(:,3);
    vsmTable{iScenario}.Rear_Steering_fault_deg = vsmTestRuns.Track(iScenario).CustomerChannels(:,4);
    vsmTable{iScenario}.Rear_Steering_slew_rate_degps = vsmTestRuns.Track(iScenario).CustomerChannels(:,5);
    vsmTable{iScenario}.Perc_of_available_Front_EMtq = vsmTestRuns.Track(iScenario).CustomerChannels(:,6);
    vsmTable{iScenario}.Perc_of_available_Rear_Emtq = vsmTestRuns.Track(iScenario).CustomerChannels(:,7);
    vsmTable{iScenario}.Slew_rate_EMtq_Nmps = vsmTestRuns.Track(iScenario).CustomerChannels(:,8);
    vsmTable{iScenario}.FL_ride_height_error_m = vsmTestRuns.Track(iScenario).CustomerChannels(:,9);
    vsmTable{iScenario}.FR_ride_height_error_m = vsmTestRuns.Track(iScenario).CustomerChannels(:,10);
    vsmTable{iScenario}.RL_ride_height_error_m = vsmTestRuns.Track(iScenario).CustomerChannels(:,11);
    vsmTable{iScenario}.RR_ride_height_error_m = vsmTestRuns.Track(iScenario).CustomerChannels(:,12);
    vsmTable{iScenario}.Slew_rate_height_mps = vsmTestRuns.Track(iScenario).CustomerChannels(:,13);
    vsmTable{iScenario}.Brk_fault = vsmTestRuns.Track(iScenario).CustomerChannels(:,14);
    vsmTable{iScenario}.Very_slow_steering_action_degps = vsmTestRuns.Track(iScenario).CustomerChannels(:,15);
    vsmTable{iScenario}.Slow_steering_action_degps = vsmTestRuns.Track(iScenario).CustomerChannels(:,16);
    vsmTable{iScenario}.Brk_applied_after_fault_perc = vsmTestRuns.Track(iScenario).CustomerChannels(:,17);
    vsmTable{iScenario}.FTTI = vsmTestRuns.Track(iScenario).CustomerChannels(:,18);
    vsmTable{iScenario}.Unique_Testrun_ID = vsmTestRuns.Track(iScenario).CustomerChannels(:,19);

    
    %Setting demand gear
    vsmTestRuns.Track(iScenario).DMDGearMap.Dist = [distance(1); distance(end)]; % distance;
    vsmTestRuns.Track(iScenario).DMDGearMap.DMDGear = -2 * ones(1, length(vsmTestRuns.Track(iScenario).DMDGearMap.Dist)); %D is -1
    vsmTable{iScenario}.demandGear = repmat("D", size(vsmTable{iScenario}.demandGear));

    %Setting DMDSpeed
    vsmTestRuns.Track(iScenario).DMDSpeedMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).DMDSpeedMap.DMDSpeed = zeros(1, length(vsmTestRuns.Track(iScenario).DMDSpeedMap.Dist));

    %Setting disable gear shift for simulation
    vsmTestRuns.Track(iScenario).DisableGSMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).DisableGSMap.DisableGS = zeros(1, length([distance(1); distance(end)]));
    vsmTable{iScenario}.disableGearShift = repmat("FALSE", size(vsmTable{iScenario}.disableGearShift));

    %Setting overall grip
    if isnan(str2double(scenarioListCells{iScenario, idx_road_friction_coefficient}))
        friction_list = strsplit(scenarioListCells{iScenario, idx_road_friction_coefficient}, '/');
        friction = str2double(friction_list(1));
        friction_ratio = str2double(friction_list(2)) / friction;
    else
        friction = str2double(scenarioListCells{iScenario, idx_road_friction_coefficient});
        friction_ratio = 1;
    end
    vsmTestRuns.Track(iScenario).GripMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).GripMap.Grip(1) = 100 * friction;
    vsmTestRuns.Track(iScenario).GripMap.Grip(2) = 100 * friction;
    vsmTable{iScenario}.gripOverall = vsmTestRuns.Track(iScenario).GripMap.Grip(1) * ones(length(distance), 1);

    %Checking and setting front left grip
    vsmTestRuns.Track(iScenario).GripFLMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).GripFLMap.Grip(1) = 100;
    vsmTestRuns.Track(iScenario).GripFLMap.Grip(2) = 100;
    vsmTable{iScenario}.gripFL = vsmTestRuns.Track(iScenario).GripFLMap.Grip(1) * ones(length(distance), 1);

    %Checking and setting front right grip   
    vsmTestRuns.Track(iScenario).GripFRMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).GripFRMap.Grip(1) = 100 * friction_ratio;
    vsmTestRuns.Track(iScenario).GripFRMap.Grip(2) = 100 * friction_ratio;
    vsmTable{iScenario}.gripFR = vsmTestRuns.Track(iScenario).GripFRMap.Grip(1) * ones(length(distance), 1);

    %Checking and setting rear left grip    
    vsmTestRuns.Track(iScenario).GripRLMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).GripRLMap.Grip(1) = 100;
    vsmTestRuns.Track(iScenario).GripRLMap.Grip(2) = 100;
    vsmTable{iScenario}.gripRL = vsmTestRuns.Track(iScenario).GripRLMap.Grip(1) * ones(length(distance), 1);

    %Checking and setting rear right grip    
    vsmTestRuns.Track(iScenario).GripRRMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).GripRRMap.Grip(1) = 100 * friction_ratio;
    vsmTestRuns.Track(iScenario).GripRRMap.Grip(2) = 100 * friction_ratio;
    vsmTable{iScenario}.gripRR = vsmTestRuns.Track(iScenario).GripRRMap.Grip(1) * ones(length(distance), 1);

    %No lat grip values needed since UseLongLatGripMaps is set to 0
    vsmTestRuns.Track(iScenario).LatGripFLMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).LatGripFRMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).LatGripRLMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).LatGripRRMap.Dist = [distance(1); distance(end)];

    %No long grip values needed since UseLongLatGripMaps is set to 0
    vsmTestRuns.Track(iScenario).LongGripFLMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).LongGripFRMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).LongGripRLMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).LongGripRRMap.Dist = [distance(1); distance(end)];

    %Setting max gear    
    vsmTestRuns.Track(iScenario).MaxGearMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).MaxGearMap.MaxGear(1) = 7;
    vsmTestRuns.Track(iScenario).MaxGearMap.MaxGear(2) = 7;
    vsmTable{iScenario}.maxGear = vsmTestRuns.Track(iScenario).MaxGearMap.MaxGear(1) * ones(length(distance), 1);

    %Setting min gear    
    vsmTestRuns.Track(iScenario).MinGearMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).MinGearMap.MinGear(1) = 1;
    vsmTestRuns.Track(iScenario).MinGearMap.MinGear(2) = 1;
    vsmTable{iScenario}.minGear = vsmTestRuns.Track(iScenario).MinGearMap.MinGear(1) * ones(length(distance), 1);

    
    %Setting banking    
    vsmTestRuns.Track(iScenario).RBMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).RBMap.Banking(1) = 0;
    vsmTestRuns.Track(iScenario).RBMap.Banking(2) = 0;
    vsmTable{iScenario}.banking = vsmTestRuns.Track(iScenario).RBMap.Banking(1) * ones(length(distance), 1);

    %Setting gradient   
    vsmTestRuns.Track(iScenario).RGMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).RGMap.RG(1) = scenarioListCells{iScenario, idx_road_gradient};
    vsmTestRuns.Track(iScenario).RGMap.RG(2) = scenarioListCells{iScenario, idx_road_gradient};
    vsmTable{iScenario}.roadGradient = vsmTestRuns.Track(iScenario).RGMap.RG(1) * ones(length(distance), 1);

    %Setting SteerAngleMap    
    vsmTestRuns.Track(iScenario).SteerAngleMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).SteerAngleMap.SteerAngle = zeros(size([distance(1); distance(end)]));

    %Setting steer mode    
    vsmTestRuns.Track(iScenario).SteerModeMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).SteerModeMap.SteerMode = zeros(1, length([distance(1); distance(end)]));
    
    %Setting VSM default values
    vsmTable{iScenario}.steerMode = repmat("Driver", size(vsmTable{iScenario}.steerMode));
    vsmTable{iScenario}.starterBit = repmat("ON", size(vsmTable{iScenario}.starterBit));
    vsmTable{iScenario}.demandTCC = repmat("UseMap", size(vsmTable{iScenario}.demandTCC));
    vsmTable{iScenario}.controlExternTorque = repmat("FALSE", size(vsmTable{iScenario}.controlExternTorque));
    vsmTable{iScenario}.driveTriggerDisabled = repmat("FALSE", size(vsmTable{iScenario}.driveTriggerDisabled));
    vsmTable{iScenario}.currentManeuver = repmat("UNKNOWN", size(vsmTable{iScenario}.currentManeuver));

    if vsmTestRuns.Track(iScenario).CycleBase == 1 %Time based simulation
        speedMax = 100;
        distanceTrack = zeros(size(distance));
        for i = 1:length(distanceTrack)
            distanceTrack(i) = (distance(i) - distance(1)) * speedMax;
        end
    else
        distanceTrack = distance;
    end
      
    %Setting distance
    vsmTestRuns.Track(iScenario).dist = distance;
    
    %Setting TrackInfo    
    vsmTestRuns.Track(iScenario).TrackInfo.Dist = distanceTrack;
    vsmTestRuns.Track(iScenario).TrackInfo.Length = distanceTrack(end);
    vsmTestRuns.Track(iScenario).TrackInfo.Name = scenarioName;
    vsmTestRuns.Track(iScenario).TrackInfo.Sectors.Pos = distanceTrack(end);
    vsmTestRuns.Track(iScenario).TrackInfo.Segments.Pos = distanceTrack(end);
    
    vsmTestRuns.Track(iScenario).TrackInfo.Dist = distanceTrack;
    vsmTestRuns.Track(iScenario).TrackInfo.Speed = vsmTestRuns.Track(iScenario).v;
    
    radiusText = Config.get_value(scenarioListCells{iScenario, idx_constant_road_radius});
    if isnumeric(radiusText)
        if radiusText > 0
            curvature = 1 / radiusText;
        else
            curvature = 0;
        end
    elseif ischar(radiusText)
        if strcmpi(radiusText, 'straight')
            curvature = 0;
        else
            try
                radiusNum = str2num(radiusText);
                if radiusNum > 0
                    curvature = 1 / radiusNum;
                else
                    curvature = 0;
                end
            catch
                error('Error: Scenario %s: Road radius (%s) is invalid.', testRunIDs{iScenario}, radiusText)
            end
        end
    else 
        error('Error: Scenario %s: Road radius (%s) is invalid.', testRunIDs{iScenario}, radiusText)
    end

    [trackX, trackY, trackZ, resultCurvature] = Config.get_track(vsmTestRuns.Track(iScenario).TrackInfo.Dist, curvature, vsmTestRuns.Track(iScenario).RGMap.RG(1), desiredSpeed);
    vsmTestRuns.Track(iScenario).TrackInfo.TrackX = trackX;
    vsmTestRuns.Track(iScenario).TrackInfo.TrackY = trackY;
    vsmTestRuns.Track(iScenario).TrackInfo.TrackZ = trackZ;
    vsmTestRuns.Track(iScenario).TrackInfo.WidthLeft = zeros(1, length(distanceTrack));
    vsmTestRuns.Track(iScenario).TrackInfo.WidthRight = zeros(1, length(distanceTrack));

    %Setting velocity limit
    vsmTestRuns.Track(iScenario).VelocityLimit.DemandSpeedMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).VelocityLimit.DemandSpeedMap.Speed = zeros(1, length([distance(1); distance(end)]));

    %Setting roughness    
    vsmTestRuns.Track(iScenario).ZSMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).ZSMap.ZS_FL = zeros(1, length([distance(1); distance(end)]));
    vsmTestRuns.Track(iScenario).ZSMap.ZS_FR = zeros(1, length([distance(1); distance(end)]));
    vsmTestRuns.Track(iScenario).ZSMap.ZS_RL = zeros(1, length([distance(1); distance(end)]));
    vsmTestRuns.Track(iScenario).ZSMap.ZS_RR = zeros(1, length([distance(1); distance(end)]));
                        
    vsmTable{iScenario}.roadRoughnessFL = vsmTestRuns.Track(iScenario).ZSMap.ZS_FL(1) * ones(length(distance), 1);
    vsmTable{iScenario}.roadRoughnessFR = vsmTestRuns.Track(iScenario).ZSMap.ZS_FR(1) * ones(length(distance), 1);
    vsmTable{iScenario}.roadRoughnessRL = vsmTestRuns.Track(iScenario).ZSMap.ZS_RL(1) * ones(length(distance), 1);
    vsmTable{iScenario}.roadRoughnessRR = vsmTestRuns.Track(iScenario).ZSMap.ZS_RR(1) * ones(length(distance), 1);

    %Setting curvature for simulation
    vsmTestRuns.Track(iScenario).k = resultCurvature;
    
    vsmTable{iScenario}.Curvature = resultCurvature';

    %Setting kNorm value
    vsmTestRuns.Track(iScenario).kNormMap.Dist = [distance(1); distance(end)];
    vsmTestRuns.Track(iScenario).kNormMap.kNorm = zeros(1,length([distance(1); distance(end)]));
    
    timeElapsed = toc;

    %fprintf('Generating Testruns: %.1f%%. Time: %.2f s, %.2f s, %.2f, %.2f, %.2f, %.2f s\n', iScenario/length(testRunIDs)*100, time1, time2, time3, time4, time5, time6);
    progress = iScenario/length(testRunIDs) * 100;
    timeRemaining = (1 / iScenario * length(testRunIDs) - 1) * timeElapsed;
    fprintf('Generating Testruns: %.1f%%. Time elapsed: %.0f s. Remaining: %.0f s\n', iScenario/length(testRunIDs)*100, timeElapsed, ceil(timeRemaining));
end

%Saving testruns into an Excel sheet
if skip_sheet_generation ~= 1
    hExcel = actxserver('Excel.Application');
    hExcel.visible = 1;
    hExcel.DisplayAlerts = false;
    workBook = hExcel.Workbooks.Add;
    workSheet = hExcel.ActiveWorkbook.Sheets;   
    try
        tic     
        for iTable = 1 : length(vsmTable)
            if iTable > 1
                workSheet.Add([], workSheet.Item(workSheet.Count));
            end
            hSheet = workSheet.Item(workSheet.Count);
            hSheet.Name = num2str(vsmTable{iTable}.Unique_Testrun_ID(1),'%05.f');
            hSheet.Activate;
            table = [vsmTable{iTable}.Properties.VariableNames; table2cell(vsmTable{iTable})];
            sizeTable = size(table);
            rowNum = sizeTable(1);
            colNum = sizeTable(2);
            eActivesheetRange = get(hExcel.Activesheet,'Range', ['A1:', Config.num2xlcol(colNum), num2str(rowNum)]);
            eActivesheetRange.Value = table;
            
            timeElapsed = toc;
            progress = iTable/length(vsmTable) * 100;
            timeRemaining = (1 / iTable * length(vsmTable) - 1) * timeElapsed;
            fprintf('Saving Testruns: %.1f%%. Time elapsed: %.0f s. Remaining: %.0f s\n', iTable/length(vsmTable)*100, timeElapsed, ceil(timeRemaining));
        end
        hSheet = workSheet.Item(1);
        hSheet.Activate;

        currentFolder = pwd;
        SaveAs(hExcel.ActiveWorkbook, append(currentFolder, '\', testrun_list_path))
        hExcel.ActiveWorkbook.Saved = 1;
        Close(workBook)
        Quit(hExcel)
        delete(hExcel)
    catch exception
        try
            Close(workBook)
            Quit(hExcel)
            delete(hExcel)
        catch
        end
        warning(getReport(exception, 'extended'))
        warning('Excel could not be started. Using Excel could speed up the saving process.')
        if exist(testrun_list_path, 'file') == 2
            delete(testrun_list_path);
        end
        for iTable = 1 : length(vsmTable)
            tic
            writetable(vsmTable{iTable}, testrun_list_path, 'Sheet', num2str(vsmTable{iTable}.Unique_Testrun_ID(1),'%05.f'))
            fprintf('Saving Testruns: %.1f%%. Time: %.2f s\n', iTable/length(vsmTable)*100, toc);
        end 
    end
end

%Saving testruns in VSM format
save(vsm_testrun_path, 'vsmTestRuns')
%save('D:\Huawei_FUSA\03_FuSa\02_Scenario Script\01_Preprocessing\Maneuvers\VSM_Testrun_test.vsd', 'vsmTestRuns')
fprintf('Done\n');