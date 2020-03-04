
import logging
logger = logging.getLogger(__name__)


def beef(number):
    print(f'this is my recursion number {number}')
    logger.critical(f'this is my recursion number {number}')
    if beef.x < 3:
        number += 2
        beef.x += 1
        beef(number)

    return number

beef.x = 0
