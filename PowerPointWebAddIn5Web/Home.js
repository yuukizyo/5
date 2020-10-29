
(function () {
    "use strict";

    var messageBanner;

    // 新しいページが読み込まれるたびに初期化関数を実行する必要があります。
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // FabricUI 通知メカニズムを初期化して、非表示にします
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            $('#get-data-from-selection').click(getDataFromSelection);
        });
    };

    // 文書の現在の選択範囲からデータを読み取り、通知を表示します
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('選択されたテキスト:', '"' + result.value + '"');
                } else {
                    showNotification('エラー:', result.error.message);
                }
            }
        );
    }

    // 通知を表示するヘルパー関数
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
