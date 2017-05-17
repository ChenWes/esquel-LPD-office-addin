var ngapp = angular.module('ngapp', ['ngMaterial', 'angularGrid']);

ngapp.controller('MainCtrl', function ($scope, $mdToast, $http, $q) {
    var vm = this;
    $scope.card = {};
    $scope.card.title = 'test';

    //page
    vm.page = 1;
    vm.pagesize = 50;

    vm.styles = [];
    vm.loadingMore = false;

    //search condition
    $scope.SearchInfo = {
        StyleID: ''
    }

    vm.doSearch = function () {
        vm.styles = [];
        vm.loadingMore = false;

        vm.loadMoreShots();
    }

    vm.loadMoreShots = function () {

        if (vm.loadingMore) return;

        // var deferred = $q.defer();
        vm.loadingMore = true;

        var promise;
        var wsUrl = 'https://dev-esb.esquel.com:6114/sprintwes/api/v1/styleproduct/' + vm.pagesize + '/' + vm.page;

        var soapRequest;


        console.log($scope.SearchInfo.StyleID);

        if ($scope.SearchInfo.StyleID != '') {
            soapRequest = {
                "filterType": "LEAF",
                "filters": [{}],
                "attributeName": "item_number",
                "searchOperator": "eq",
                "filterValue": $scope.SearchInfo.StyleID
            };

            vm.page = 1;
        } else {
            soapRequest = {};            
        }

        $.ajax({
            type: "POST",
            url: wsUrl,
            Connection: "Keep-Alive",
            processData: true,
            async: false,
            data: JSON.stringify(soapRequest),
            headers: {
                "Content-Type": "application/json",
                "Accept": "application/json"
            },
            success: function (data, status, req) {
                $("#user").text("Status is definitely:" + status);
                var shotsTmp = angular.copy(vm.shots);

                var responsedata = data;

                if (responsedata.resultType === "SUCCESS") {
                    if (responsedata.results) {
                        shotsTmp = responsedata.results[0].data;

                        vm.styles = shotsTmp;
                        if (!vm.styles[0]) {
                            showError('No Style Mathced.');
                        } else {
                            showInfo('We Found Page [' + vm.page + '] ' + vm.styles.length + ' Style.')
                            vm.page++;
                        }
                    }
                    else {
                        showError('WebAPI Return Error:' + responsedata.resultMsg);
                    }
                } else {
                    showError('WebAPI Return Error:' + responsedata.resultMsg);
                }

                vm.loadingMore = false;
            },
            error: function (jqXHR, textStatus, errorThrown) {
                vm.loadingMore = false;

                showError('Call WebAPI Have Error.');
            },
        }).fail(function (jqXHR, textStatus, errorThrown) {
            vm.loadingMore = false;

            showError('Call WebAPI Have Error.');
        });

        return promise;
    };


    showError = function (msg) {
        $mdToast.show(
            $mdToast.simple()
                .textContent(msg)
                .hideDelay(5000)
        );
    }

    showInfo = function (info) {
        $mdToast.show(
            $mdToast.simple()
                .textContent(info)
                .hideDelay(5000)
        );
    }

    //vm.loadMoreShots();
});


ngapp.filter('unsafe', function ($sce) { return $sce.trustAsHtml; });