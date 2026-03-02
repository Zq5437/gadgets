"""
PDF 批量解密工具 - 提供美观的终端交互界面
"""

import sys
from pathlib import Path
from typing import Optional, List, Callable
from dataclasses import dataclass
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

import pikepdf
from rich.console import Console
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn, TimeRemainingColumn
from rich.table import Table
from rich import box
from rich.prompt import Prompt, Confirm, IntPrompt
from rich.live import Live
from rich.layout import Layout
from rich.text import Text
from rich.align import Align

console = Console()


@dataclass
class DecryptResult:
    """解密结果数据类"""
    filename: str
    status: str  # 'success', 'error', 'skipped'
    message: str
    output_path: Optional[Path] = None
    error_type: Optional[str] = None


class PDFBatchDecryptor:
    """PDF 批量解密器 - 支持美观的终端交互"""

    def __init__(self, console: Optional[Console] = None):
        self.console = console or Console()
        self.results: List[DecryptResult] = []
        self._lock = threading.Lock()

    def show_banner(self):
        """显示程序横幅"""
        banner = Panel.fit(
            "[bold cyan]PDF Batch Decryptor[/bold cyan]\n"
            "[dim]安全高效的 PDF 批量解密工具[/dim]",
            border_style="cyan",
            box=box.ROUNDED
        )
        self.console.print(banner)

    def get_input_interactive(self) -> tuple:
        """交互式获取输入参数"""
        self.console.print("\n[bold yellow]⚙️  配置参数[/bold yellow]\n")

        # 输入目录
        while True:
            input_dir = Prompt.ask("📁 输入目录 (包含加密 PDF 的文件夹)")
            input_dir = input_dir.strip().strip('\'"')
            input_path = Path(input_dir).expanduser()
            if input_path.exists() and input_path.is_dir():
                break
            self.console.print("[red]❌ 目录不存在，请重新输入[/red]")

        # 检查是否有 PDF 文件
        pdf_files = list(input_path.glob("*.pdf"))
        if not pdf_files:
            self.console.print(
                f"[yellow]⚠️  在 {input_dir} 中未找到 PDF 文件[/yellow]")
            if not Confirm.ask("是否继续?"):
                sys.exit(0)
        else:
            self.console.print(
                f"[green]✓ 找到 {len(pdf_files)} 个 PDF 文件[/green]")

        # 输出目录
        default_output = str(input_path.parent / "decrypted_pdfs")
        output_dir = Prompt.ask("📁 输出目录", default=default_output)
        output_path = Path(output_dir).expanduser()

        # 密码
        password = Prompt.ask("🔑 PDF 密码", password=True)

        # 密码库选项
        use_password_list = Confirm.ask("是否使用密码库（密码列表）?", default=True)
        password_list = []
        if use_password_list:
            while True:
                default_output = "./password_vault.txt"
                list_file = Prompt.ask("📁 密码库文件路径", default=default_output)
                list_file = list_file.strip().strip('\'"')
                list_path = Path(list_file).expanduser()
                if list_path.exists() and list_path.is_file():
                    try:
                        with open(list_path, 'r', encoding='utf-8') as f:
                            # 读取非空行，去除首尾空白
                            password_list = [line.strip()
                                             for line in f if line.strip()]
                        if not password_list:
                            self.console.print("[yellow]⚠️  密码库文件为空[/yellow]")
                            if not Confirm.ask("是否继续?"):
                                continue
                        break
                    except Exception as e:
                        self.console.print(f"[red]❌ 读取密码库失败: {e}[/red]")
                        if not Confirm.ask("是否重新输入?"):
                            break
                else:
                    self.console.print("[red]❌ 文件不存在，请重新输入[/red]")
        # 合并用户密码和密码库（去重，用户密码优先）
        passwords = [password] + [p for p in password_list if p != password]

        # 高级选项
        self.console.print("\n[bold yellow]🔧 高级选项[/bold yellow]")
        use_threads = Confirm.ask("是否启用多线程处理?", default=True)
        max_workers = 4
        if use_threads:
            max_workers = IntPrompt.ask("并发线程数", default=4)

        prefix = Prompt.ask("输出文件名前缀", default="decrypted_")

        return {
            'input_dir': input_path,
            'output_dir': output_path,
            'password': password,          # 保留原字段（可选）
            'passwords': passwords,        # 新增密码列表
            'use_threads': use_threads,
            'max_workers': max_workers,
            'prefix': prefix
        }

    def decrypt_single(self, pdf_file: Path, output_dir: Path, passwords: List[str], prefix: str) -> DecryptResult:
        """解密单个 PDF 文件，尝试多个密码"""
        last_error = None
        for password in passwords:
            try:
                with pikepdf.open(pdf_file, password=password) as pdf:
                    output_file = output_dir / f"{prefix}{pdf_file.name}"

                    # 检查文件是否已存在
                    if output_file.exists():
                        return DecryptResult(
                            filename=pdf_file.name,
                            status='skipped',
                            message=f"输出文件已存在: {output_file.name}",
                            output_path=output_file
                        )

                    pdf.save(output_file)
                    return DecryptResult(
                        filename=pdf_file.name,
                        status='success',
                        message="解密成功",
                        output_path=output_file
                    )

            except pikepdf.PasswordError:
                # 密码错误，继续尝试下一个
                last_error = pikepdf.PasswordError
                continue
            except pikepdf.PdfError as e:
                # PDF格式错误，无需继续尝试
                return DecryptResult(
                    filename=pdf_file.name,
                    status='error',
                    message=f"PDF 格式错误: {str(e)}",
                    error_type='PdfError'
                )
            except Exception as e:
                # 其他未知错误，提前终止
                return DecryptResult(
                    filename=pdf_file.name,
                    status='error',
                    message=f"未知错误: {str(e)}",
                    error_type=type(e).__name__
                )
        # 所有密码尝试失败
        return DecryptResult(
            filename=pdf_file.name,
            status='error',
            message="密码错误（尝试所有密码均失败）",
            error_type='PasswordError'
        )

    def process_batch(self, config: dict, progress_callback: Optional[Callable] = None):
        """批量处理 PDF 文件"""
        input_dir = config['input_dir']
        output_dir = config['output_dir']
        passwords = config.get('passwords', [config['password']])  # 兼容旧配置
        use_threads = config.get('use_threads', True)
        max_workers = config.get('max_workers', 4)
        prefix = config.get('prefix', 'decrypted_')

        # 创建输出目录
        output_dir.mkdir(parents=True, exist_ok=True)

        # 获取所有 PDF 文件
        pdf_files = list(input_dir.glob("*.pdf"))
        total = len(pdf_files)

        if total == 0:
            self.console.print("[yellow]⚠️  未找到 PDF 文件[/yellow]")
            return

        self.results = []

        if use_threads and total > 1:
            self._process_with_threads(
                pdf_files, output_dir, passwords, prefix, max_workers)  # 传递 passwords
        else:
            self._process_sequential(pdf_files, output_dir, passwords, prefix)  # 传递 passwords

    def _process_with_threads(self, pdf_files: List[Path], output_dir: Path, passwords: str, prefix: str, max_workers: int):
        """使用多线程处理"""
        total = len(pdf_files)
        completed = 0

        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(bar_width=40),
            TaskProgressColumn(),
            TimeRemainingColumn(),
            console=self.console
        ) as progress:
            task = progress.add_task(f"[cyan]解密中...", total=total)

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_file = {
                    executor.submit(self.decrypt_single, pdf_file, output_dir, passwords, prefix): pdf_file
                    for pdf_file in pdf_files
                }

                for future in as_completed(future_to_file):
                    result = future.result()
                    with self._lock:
                        self.results.append(result)
                        completed += 1
                        progress.update(
                            task, advance=1, description=f"[cyan]处理中... [{completed}/{total}] {result.filename[:30]}")

    def _process_sequential(self, pdf_files: List[Path], output_dir: Path, passwords: str, prefix: str):
        """顺序处理"""
        total = len(pdf_files)

        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(bar_width=40),
            TaskProgressColumn(),
            TimeRemainingColumn(),
            console=self.console
        ) as progress:
            task = progress.add_task("[cyan]解密中...", total=total)

            for pdf_file in pdf_files:
                result = self.decrypt_single(pdf_file, output_dir, passwords, prefix)
                self.results.append(result)
                progress.update(
                    task, advance=1, description=f"[cyan]处理中... {pdf_file.name[:30]}")

    def show_results(self):
        """显示处理结果表格"""
        if not self.results:
            return

        # 统计信息
        success_count = sum(1 for r in self.results if r.status == 'success')
        error_count = sum(1 for r in self.results if r.status == 'error')
        skip_count = sum(1 for r in self.results if r.status == 'skipped')

        # 创建结果表格
        table = Table(
            title="📊 处理结果",
            box=box.ROUNDED,
            show_header=True,
            header_style="bold magenta"
        )
        table.add_column("文件名", style="cyan", no_wrap=True, width=30)
        table.add_column("状态", justify="center", width=10)
        table.add_column("信息", style="dim", width=40)

        for result in self.results:
            if result.status == 'success':
                status_text = "[green]✓[/green]"
                style = "green"
            elif result.status == 'error':
                status_text = "[red]✗[/red]"
                style = "red"
            else:
                status_text = "[yellow]⊘[/yellow]"
                style = "yellow"

            table.add_row(
                result.filename[:28] +
                "..." if len(result.filename) > 30 else result.filename,
                status_text,
                result.message,
                style=style
            )

        self.console.print(table)

        # 统计面板
        stats = Panel(
            f"[bold]总计:[/bold] {len(self.results)} | "
            f"[green]成功: {success_count}[/green] | "
            f"[red]失败: {error_count}[/red] | "
            f"[yellow]跳过: {skip_count}[/yellow]",
            title="统计",
            border_style="blue"
        )
        self.console.print(stats)

    def export_results(self, output_file: Optional[Path] = None):
        """导出结果到文本文件"""
        if not self.results:
            return

        if output_file is None:
            output_file = Path("decrypt_report.txt")

        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("PDF 解密报告\n")
            f.write("=" * 50 + "\n\n")
            for result in self.results:
                f.write(f"文件: {result.filename}\n")
                f.write(f"状态: {result.status}\n")
                f.write(f"信息: {result.message}\n")
                if result.output_path:
                    f.write(f"输出: {result.output_path}\n")
                f.write("-" * 50 + "\n")

        self.console.print(f"[dim]📄 报告已保存到: {output_file}[/dim]")


def main():
    """主函数 - 交互式模式"""
    decryptor = PDFBatchDecryptor()
    decryptor.show_banner()

    # 获取配置
    config = decryptor.get_input_interactive()

    # 确认开始
    console.print("\n[bold]配置摘要:[/bold]")
    console.print(f"  输入: {config['input_dir']}")
    console.print(f"  输出: {config['output_dir']}")
    console.print(
        f"  线程: {config['max_workers'] if config['use_threads'] else '单线程'}")

    if not Confirm.ask("\n是否开始处理?", default=True):
        console.print("[yellow]已取消[/yellow]")
        return

    # 开始处理
    console.print("\n[bold green]🚀 开始处理...[/bold green]\n")
    decryptor.process_batch(config)

    # 显示结果
    decryptor.show_results()

    # 询问是否导出报告
    if Confirm.ask("\n是否导出详细报告?"):
        decryptor.export_results()

    console.print("\n[bold green]✨ 处理完成![/bold green]")


if __name__ == "__main__":
    main()
