﻿var ngapp = angular.module('ngapp', ['ngMaterial', 'angularGrid']);

ngapp.controller('MainCtrl', function ($scope, $http, $q) {
    var vm = this;
    $scope.card = {};
    $scope.card.title = 'test';
    vm.page = 1;
    vm.pagesize = 100;
    vm.searchYear = 2016;
    vm.shots = [];
    vm.loadingMore = false;

    //search condition
    $scope.SearchInfo = {
        Year: "2016"
    }
    vm.doSearch = function () {
        vm.page = 0;
        vm.shots = [];
        vm.loadingMore = false;
        vm.searchYear = $scope.SearchInfo.Year;
        if (vm.searchYear == "") {
            vm.searchYear = "2016";
        }
        vm.loadMoreShots();
    }

    vm.loadMoreShots = function () {

        if (vm.loadingMore) return;
        vm.page++;
        // var deferred = $q.defer();
        vm.loadingMore = true;

        var promise;
        var wsUrl = 'https://dev-esb.esquel.com:6114/ESB/RequestReply';
        var soapRequest = '<nam:ArrayOfSelectionFilter xmlns:nam="http://www.esquel.com/Product/namespace/"><nam:SelectionFilter><nam:AttributeName>cn_year</nam:AttributeName><nam:FilterType>LEAF</nam:FilterType><nam:FilterValue>' + vm.searchYear + '</nam:FilterValue><nam:SearchOperator>EQ</nam:SearchOperator></nam:SelectionFilter><nam:pageIndex>' + vm.page + '</nam:pageIndex><nam:pageSize>' + vm.pagesize + '</nam:pageSize></nam:ArrayOfSelectionFilter>';
        $.ajax({
            type: "POST",
            url: wsUrl,
            contentType: "text/xml;charset=UTF-8",
            dataType: "text",
            Connection: "Keep-Alive",
            processData: true,
            async: false,
            data: soapRequest,
            headers: {
                "Jms-Destination": "cn.grm.instantNoodle.getFabricPartList",
                "Jms-Timeout": 100
            },
            success: function (data, status, req) {
                $("#user").text("Status is definitely:" + status);
                var shotsTmp = angular.copy(vm.shots);

                //shotsTmp = shotsTmp.concat($(data).find("Woven"));
                var x2js = new X2JS();
                var aftCnv = x2js.xml_str2json(data);

                promise = aftCnv.ESBWrapperOutput.ResponseObject.ResponseData.Data.FabricItems.Woven;
                shotsTmp = shotsTmp.concat(promise);
                vm.shots = shotsTmp;
                vm.loadingMore = false;
            },
            error: function (jqXHR, textStatus, errorThrown) {
                vm.loadingMore = false;
            },
        }).fail(function (jqXHR, textStatus, errorThrown) {
            vm.loadingMore = false;
        });

        return promise;
    };

    //vm.loadMoreShots();

});
ngapp.filter('unsafe', function ($sce) { return $sce.trustAsHtml; });

function loadJSON(callback) {

    var xobj = new XMLHttpRequest();
    xobj.overrideMimeType("application/json");
    xobj.open('GET', 'photo.json', true); // Replace 'my_data' with the path to your file
    xobj.onreadystatechange = function () {
        if (xobj.readyState == 4 && xobj.status == "200") {
            // Required use of an anonymous callback as .open will NOT return a value but simply returns undefined in asynchronous mode
            callback(xobj.responseText);
        }
    };
    xobj.send(null);
}