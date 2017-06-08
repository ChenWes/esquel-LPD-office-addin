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

    }]);;

ngapp.controller('MainCtrl', ['$scope', '$mdToast', '$http', 'adalAuthenticationService', '$q', function ($scope, $mdToast, $http, adalService, $q) {
    var vm = this;

    $scope.login = function () {
        adalService.login();
    };
    $scope.logout = function () {
        adalService.logOut();
    };

    $scope.card = {};
    $scope.card.title = 'Esquel';

    $scope.attrCSS = 'float: left; padding: 5px; border: 1px solid #EEEEEE; background-color: #EEEEEE;margin:5px';

    //$scope.defaults.gridNo = 1;  
    vm.page = 1;
    vm.pagesize = 20;
    vm.styles = [];

    //just flag , prevent mutil call
    vm.loadingMore = false;

    //search condition
    $scope.SearchInfo = {
        styleID: ''
    }


    /* button search function */
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

    }

    /* copy content to mail */
    vm.doCopyContent = function (dvId) {

        //var item = vm.shots[dvId];
        var ct = $("#dv_" + dvId);

        var item = Office.cast.item.toItemCompose(Office.context.mailbox.item);

        item.body.getTypeAsync(
            function (result) {
                if (result.status != Office.AsyncResultStatus.Failed) {
                    // Successfully got the type of item body.
                    // Set data of the appropriate type in body.
                    if (result.value == Office.MailboxEnums.BodyType.Html) {
                        // Body is of HTML type.
                        // Specify HTML in the coercionType parameter
                        // of setSelectedDataAsync.
                        item.body.setSelectedDataAsync(
                            ct.html(),
                            {
                                coercionType: Office.CoercionType.Html,
                                asyncContext: { var3: 1, var4: 2 }
                            },
                            function (asyncResult) {
                                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                                    item.subject.setAsync(asyncResult.error.message);
                                }
                            });
                    }
                    else {
                        // Body is of text type. 
                        item.body.setSelectedDataAsync(
                            ct.html(),
                            {
                                coercionType: Office.CoercionType.Text,
                                asyncContext: { var3: 1, var4: 2 }
                            },
                            function (asyncResult) {
                                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                                    item.subject.setAsync(asyncResult.error.message);
                                }
                            });
                    }
                }
            });
    }

    /* call WebAPI function */
    vm.loadMoreShots = function (token) {

        if (vm.loadingMore) return;

        // var deferred = $q.defer();
        vm.loadingMore = true;

        var promise;
        var wsUrl = 'https://designer-workbench-dev.azurewebsites.net/api/v1/styleproduct/' + vm.pagesize + '/' + vm.page;
        var soapRequest = '';

        if ($scope.SearchInfo.styleID != '') {
            soapRequest = {
                "filterType": "LEAF",
                "filters": [{}],
                "attributeName": "item_number",
                "searchOperator": "eq",
                "filterValue": $scope.SearchInfo.styleID
            };

            vm.page = 1;
        } else {
            //search all result.
            soapRequest = {};
        }

        console.log('token------------');
        console.log(token);

        var req = {
            method: 'POST',
            url: wsUrl,
            data: JSON.stringify(soapRequest),
            headers: {
                "Content-Type": "application/json",
                "Accept": "application/json",
                "Authorization": "Bearer " + token
            },
            cache: false
        }

        var doneFun = function (resp) {
            var shotsTmp = angular.copy(vm.shots);

            if (resp.data) {
                var responsedata = resp.data;

                if (responsedata.resultType === "SUCCESS") {
                    if (responsedata.results) {
                        shotsTmp = responsedata.results[0].data;
                        var totalPage = responsedata.results[0].totalPage;

                        vm.styles = shotsTmp;
                        if (!vm.styles[0]) {
                            showError('No Style Mathced.');
                        } else {
                            showInfo('We Found Page [' + vm.page + '] ' + vm.styles.length + ' Style.' + token);

                            //check next page
                            if (vm.page < totalPage) {
                                vm.page++;
                            }
                        }
                    }
                    else {
                        showError('WebAPI Return Error:' + tempdata.resultMsg);
                    }
                } else {
                    showError('WebAPI Return Error:' + tempdata.resultMsg);
                }
            } else {
                showError('WebAPI Return Data Is Null Or Empty.');
            }

            vm.loadingMore = false;
        };

        var failedFun = function () {
            vm.loadingMore = false;
            showError('Call WebAPI Have Error.');
        };

        return $http(req).then(doneFun, failedFun);
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
}]);


ngapp.filter('unsafe', function ($sce) { return $sce.trustAsHtml; });

