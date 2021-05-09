import discord
from discord.ext import commands
import current_price.trade as cp

profit = 0
stock_channel_id = 835453371062288404
prefix = "!"
current = cp.CurrentPrice()
bot = commands.Bot(command_prefix=prefix)


@bot.event
async def on_ready():
    # bot의 상태
    game = discord.Game("deep learning")
    await bot.change_presence(status=discord.Status.online, activity=game)
    await bot.get_channel(stock_channel_id).send("I'm ready")   # 디스코드입장시 지정한 채널에 알림
    print("READY")


@bot.event
async def on_message(message):      # 메세지 수신
    if message.author.bot:          # 송신자가 봇이면 무시
        return None
    await bot.process_commands(message)


@bot.command(aliases=['명령어', '도움말', '도움', '헬프'])
async def react_help(ctx):
    embed = discord.Embed(title="명령어 목록", description="모든 명령어 앞에는 !를 붙여주세요.\n모의투자 대회 기능을 구현할 예정입니다.",
                          color=0x62c1cc)  # Embed의 기본 틀(색상, 메인 제목, 설명)을 잡아줍니다
    embed.add_field(name="가격 종목명", value="입력한 회사의 현재 가격이 출력됩니다", inline=False)
    embed.add_field(name="매수 종목명 수량", value="입력한 종목의 주식을 입력한 수량만큼 매수합니다", inline=False)
    embed.add_field(name="매도 종목명 수량", value="입력한 종목의 주식을 입력한 수량만큼 매도합니다", inline=False)
    embed.add_field(name="잔고", value="현재 계좌의 잔고를 출력합니다. +2 예수금, 보유중인 주식의 정보", inline=False)

    await ctx.message.channel.send(embed=embed)  # embed를 포함 한 채로 메시지를 전송합니다.
    return None


@bot.command(name="가격")
async def react_price(ctx, *args):
    if len(args) == 0:  # '!가격' 만 입력
        await ctx.channel.send('회사명 입력해 주세요.')
        return None
    print(args)
    price = current.get_current_price(str(args[0]).upper())
    if price is None:
        await ctx.channel.send('회사명을 정확히 입력해 주세요.')
    else:
        await ctx.channel.send('현재 {}의 가격은 {}원 입니다.'.format(str(args[0]).upper(), price))
    return None


@bot.command(name="매수")
async def react_buy(ctx, *args):
    if len(args) < 2:
        await ctx.channel.send('매수할 종목과 수량을 입력하세요')
        return None

    result = current.trade('2', str(args[0]).upper(), args[1])
    if result == 0:
        await ctx.channel.send('{}을(를) {}주 매수했습니다.'.format(str(args[0]).upper(), args[1]))
    elif result == -1:
        await ctx.channel.send('장중 시간이 아닙니다')
    else:
        await ctx.channel.send('매수할 종목과 수량을 정확히 입력해 주세요')
    return None


@bot.command(name="매도")
async def react_sell(ctx, *args):
    if len(args) < 2:
        await ctx.channel.send('매도할 종목과 수량을 입력하세요')
        return None

    result = current.trade('1', str(args[0]).upper(), args[1])
    if result == 0:
        await ctx.channel.send('{}을(를) {}주 매도했습니다.'.format(str(args[0]).upper(), args[1]))
    elif result == -1:
        await ctx.channel.send('장중 시간이 아닙니다')
    else:
        await ctx.channel.send('매수할 종목과 수량을 정확히 입력해 주세요')
    return None


@bot.command(name="잔고")
async def react_portfolio(ctx):
    # Send a message to channel what user sent
    result = current.get_portfolio()

    if len(result) == 0:
        await ctx.channel.send('등록된 계좌가 없습니다.')
    else:
        embed = discord.Embed(title="계좌 잔고", description="+2 예수금 : {}원".format(result[0]),
                              color=0xffffff)  # Embed의 기본 틀(색상, 메인 제목, 설명)을 잡아줍니다
        for idx, stock in enumerate(result):
            if idx == 0:
                continue
            embed.add_field(name="{}".format(stock['name']), value="수익률 : {:.3f}%\n {}주\n 현재가 : {}\n총액 : {}원"
                            .format(stock['profit'], stock['quantity'], stock['cprice'], stock['quantity'] * stock['cprice']), inline=False)

        await ctx.message.channel.send(embed=embed)  # embed를 포함 한 채로 메시지를 전송합니다.
    return None


@bot.event
async def on_command_error(ctx, error):     # 명령어 입력 에러
    embed = discord.Embed(title="명령어를 확인해 주세요",
                          color=0x62c1cc)  # Embed의 기본 틀(색상, 메인 제목, 설명)을 잡아줍니다
    embed.add_field(name="가격 종목명", value="입력한 회사의 현재 가격이 출력됩니다", inline=False)
    embed.add_field(name="매수 종목명 수량", value="입력한 종목의 주식을 입력한 수량만큼 매수합니다", inline=False)
    embed.add_field(name="매도 종목명 수량", value="입력한 종목의 주식을 입력한 수량만큼 매도합니다", inline=False)
    embed.add_field(name="잔고", value="현재 계좌의 잔고를 출력합니다. +2 예수금, 보유중인 주식의 정보", inline=False)

    await ctx.message.channel.send(embed=embed)  # embed를 포함 한 채로 메시지를 전송합니다.
    return None


def start_bot():
    bot.run("bot_token")


if __name__ == "__main__":
    start_bot()

