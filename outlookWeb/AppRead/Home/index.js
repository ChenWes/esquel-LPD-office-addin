var ngapp = angular.module('ngapp', ['ngMaterial', 'angularGrid', 'AdalAngular'])
    .config(['$httpProvider', 'adalAuthenticationServiceProvider', function ($httpProvider, adalProvider) {
        adalProvider.init(
            {
                instance: 'https://login.microsoftonline.com/',
                tenant: '29abf16e-95a2-4d13-8d51-6db1b775d45b',
                clientId: '741a869c-ce4c-46c0-8794-60a6391293ca',
                extraQueryParameter: 'nux=1',
                cacheLocation: 'localStorage' // enable this for IE, as sessionStorage does not work for localhost.
            },
            $httpProvider
        );

    }]);

ngapp.controller('MainCtrl', ['$scope', '$mdToast', '$http', 'adalAuthenticationService', '$q', function ($scope, $mdToast, $http, adalService, $q) {
    var vm = this;

    $scope.card = {};
    $scope.card.title = 'Esquel';

    $scope.login = function () {
        adalService.login();
    };
    $scope.logout = function () {
        adalService.logOut();
    };


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

        if (adalService.userInfo) {
            var resource = 'https://esquel.onmicrosoft.com/705cadd7-d8b2-44f7-9c28-3841c112f04b';
            var oldtoken = adalService.getCachedToken(resource);
            if (!oldtoken) {
                adalService.acquireToken(resource).then(function (token) {
                    vm.loadMoreShots(token);
                    $(".angular-grid-item").css("position", "relative");
                }).catch(function (err) {
                    showError(err);
                });
            }
            else {
                vm.loadMoreShots(oldtoken);
                $(".angular-grid-item").css("position", "relative");
            }
        } else {
            showError('no userinfo ,please login first');
        }

        // vm.loadMoreShots();
    }

    vm.loadMoreShots = function (token) {

        if (vm.loadingMore) return;

        // var deferred = $q.defer();
        vm.loadingMore = true;

        var promise;
        var wsUrl = 'https://designer-workbench-dev.azurewebsites.net/api/v1/styleproduct/' + vm.pagesize + '/' + vm.page;

        var soapRequest;


        // console.log($scope.SearchInfo.StyleID);

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
                "Accept": "application/json",
                "Authorization": "Bearer " + token
            },
            success: function (data, status, req) {
                $("#user").text("Status is definitely:" + status);
                var shotsTmp = angular.copy(vm.shots);

                var responsedata = data;

                if (responsedata.resultType === "SUCCESS") {
                    if (responsedata.results) {
                        shotsTmp = responsedata.results[0].data;
                        var totalPage = responsedata.results[0].totalPage;

                        vm.styles = shotsTmp;
                        if (!vm.styles[0]) {
                            showError('No Style Mathced.');
                        } else {
                            showInfo('We Found Page [' + vm.page + '] ' + vm.styles.length + ' Style.')

                            //check next page
                            if (vm.page < totalPage) {
                                vm.page++;
                            }
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
}]);


ngapp.filter('unsafe', function ($sce) { return $sce.trustAsHtml; });